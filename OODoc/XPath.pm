#-----------------------------------------------------------------------------
#
#	$Id : XPath.pm 1.115 2004-08-03 JMG$
#
#	Initial developer: Jean-Marie Gouarne
#	Copyright 2004 by Genicorp, S.A. (www.genicorp.com)
#	Licensing conditions:
#		- Licence Publique Generale Genicorp v1.0
#		- GNU Lesser General Public License v2.1
#	Contact: oodoc@genicorp.com
#
#-----------------------------------------------------------------------------

package	OpenOffice::OODoc::XPath;
use	5.008_000;
our	$VERSION	= 1.115;
use	XML::XPath	1.13;
use	Encode;

#------------------------------------------------------------------------------

our %XMLNAMES	=			# OODoc root element names
	(
	'meta'		=> 'office:document-meta',
	'content'	=> 'office:document-content',
	'styles'	=> 'office:document-styles',
	'settings'	=> 'office:document-settings'
	);

					# characters to be escaped in XML
our	$CHARS_TO_ESCAPE	= "\"<>'&";
					# standard external character set
our	$LOCAL_CHARSET		= 'iso-8859-1';
					# OpenOffice.org character set
our	$OO_CHARSET		= 'utf8';

#------------------------------------------------------------------------------
# basic conversion between internal & printable encodings

sub	OpenOffice::OODoc::XPath::decode_text
	{
	return Encode::encode($LOCAL_CHARSET, shift);
	}

sub	OpenOffice::OODoc::XPath::encode_text
	{
	return Encode::decode($LOCAL_CHARSET, shift);
	}

#------------------------------------------------------------------------------
# replace toString method from XML::XPath for Text and Attribute nodes

sub	XML::XPath::Node::Text::toString
	{
	my $node	= shift;

	return	XML::XPath::Node::XMLescape
			(
			Encode::encode($OO_CHARSET, $node->getNodeValue),
			$CHARS_TO_ESCAPE
			);
	}

sub	XML::XPath::Node::Attribute::toString
	{
	my $node	= shift;

	return	' '							.
		OpenOffice::OODoc::XPath::encode_text($node->getName)	.
		'="'							.
		XML::XPath::Node::XMLescape
			(
			Encode::encode($OO_CHARSET, $node->getNodeValue),
			$CHARS_TO_ESCAPE
			)						.
		'"';
	}

#------------------------------------------------------------------------------
# common search/replace text processing routine (class method)
# if $replace is a user-provided routine, it's called back with
# the current argument stack, plus the substring found

sub	OpenOffice::OODoc::XPath::_find_text
	{
	my $text	= shift;
	my $pattern	= OpenOffice::OODoc::XPath::encode_text(shift);
	my $replace	= shift;

	if (defined $pattern)
	    {
	    if (defined $replace)
		{
		if (ref $replace)
		    {
		    if ((ref $replace) eq 'CODE')
		    	{
			return undef
			  unless
			    (
			    $text =~
			    	s/($pattern)/
				    	{
					my $found = $1;
					Encode::_utf8_on($found)
						if Encode::is_utf8($text);
					my $result = &$replace(@_, $found);
					$result = $found
						unless (defined $result);
					$result;
					}
				/eg
			    );
			}
		    else
		    	{
			return undef unless ($text =~ /$pattern/);
			}
		    }
		else
		    {
		    my $r = OpenOffice::OODoc::XPath::encode_text($replace);
		    return undef unless ($text =~ s/$pattern/$r/g);
		    }
		}
	    else
		{
		return undef unless ($text =~ /$pattern/);
		}
	    }
	return $text;
	}

#------------------------------------------------------------------------------
# remove all the children of a given element (extends XML::XPath)

sub	XML::XPath::Node::Element::removeChildNodes
	{
	my $element	= shift;

	$element->removeChild($_) for $element->getChildNodes;
	}

#------------------------------------------------------------------------------
# recursive text search & replace processing (extends XML::XPath)

sub	XML::XPath::Node::Element::_search_content
	{
	my $element	= shift;
	my $content	= undef;
	foreach my $child ($element->getChildNodes)
		{
		my $text = undef;
		if	($child->isTextNode)
			{
			$text = OpenOffice::OODoc::XPath::_find_text
				($child->string_value, @_);
			if (defined $text)
				{
				$child->setNodeValue($text);
				}
			}
		elsif	($child->isElementNode)
			{
			my $t = $child->_search_content(@_);
			$text .= $t if (defined $t);
			}
		$content .= $text if (defined $text);
		}
	return $content;
	}

#------------------------------------------------------------------------------
# reserved properties (to be implemented)

sub	isCalcDocument		{}
sub	isImpressDocument	{}
sub	isDrawDocument		{}

#------------------------------------------------------------------------------
# constructor; accepts one from 3 types of parameters to create an instance:
#	file	=> a regular OpenOffice.org filename
#	archive	=> an OODoc::File, previously created object
#	xml	=> an XML string, representing an OO XML member
# if 'file' or 'archive' (not 'xml') is provided, another parameter 'member'
# must be provided in addition
# 	member	=> member of the zip archive (meta.xml, content.xml, ...)

sub	new
	{
	my $caller	= shift;
	my $class	= ref($caller) || $caller;
	my $self	=
		{
		body_path		=> '//office:body',
		auto_style_path		=> '//office:automatic-styles',
		master_style_path	=> '//office:master-styles',
		named_style_path	=> '//office:styles',
		@_
		};
	if (($self->{'file'}) && (! $self->{'archive'}))
		{
		require OpenOffice::OODoc::File;

		$self->{'archive'} = OpenOffice::OODoc::File->new
				(
				$self->{'file'},
				create		=> $self->{'create'},
				template_path	=> $self->{'template_path'}
				);
		}

	unless ($self->{'xml'})
		{
		if ($self->{'archive'})
			{
			$self->{'member'} = 'content' unless $self->{'member'};
			$self->{'xml'} =
			    $self->{'archive'}->link($self);

			unless ($self->{'element'})
				{
				my $m	= lc $self->{'member'};
				$m	=~ /(^.*)\..*/;
				$m	= $1	if $1;
				$self->{'element'} =
					$OpenOffice::OODoc::XPath::XMLNAMES{$m};
				}
			}
		else
			{
			warn "[" . __PACKAGE__ . "] No oo_archive\n";
			return undef;
			}
		}
		
	if ($self->{'element'})
		{
		my $t	= $self->{'xml'};
		my $b	= $self->{'element'};
		$t	=~ /(.*)(<\s*$b\s.*<\s*\/$b\s*>)(.*)/s;
		$self->{'begin'}	= $1; chomp $self->{'begin'};
		$self->{'xml'}		= $2; chomp $self->{'xml'};
		$self->{'end'}		= $3;
		}
		
	if ($self->{'xml'})
		{
		unless ($self->{'parser'})
			{
			if
			    (
			    $main::XML_PARSER
				&&
			    $main::XML_PARSER->isa('XML::XPath::XMLParser')
			    )
			    {
			    $self->{'parser'} = $main::XML_PARSER;
			    }
			else
			    {
			    $self->{'parser'} = XML::XPath::XMLParser->new;
			    }
			}
		$self->{'xpath'} = $self->{'parser'}->parse($self->{'xml'});
		$self->{'xml'} = undef;
		}
	else
		{
		warn "[" . __PACKAGE__ . "] No XML content\n";
		return undef;
		}
	return bless $self, $class;
	}

#------------------------------------------------------------------------------
# get a reference to the embedded XML parser for share

sub	getXMLParser
	{
	my $self	= shift;
	return $self->{'parser'};
	}

#------------------------------------------------------------------------------
# make the changes persistent in an OpenOffice.org file 

sub	save
	{
	my $self	= shift;
	my $target	= shift;

	my $filename	= ($target) ? $target : $self->{'file'};
	my $archive	= $self->{'archive'};
	unless ($archive)
		{
		warn "[" . __PACKAGE__ . "::save] No archive object\n";
		return undef;
		}
	$filename	= $archive->{'source_file'}	unless $filename;
	unless ($filename)
		{
		warn "[" . __PACKAGE__ . "::save] No target file\n";
		return undef;
		}
	my $member	= $self->{'member'};
	unless ($member)
		{
		warn "[" . __PACKAGE__ . "::save] No member\n";
		return undef;
		}
	
	my $result = $archive->save($filename);

	return $result;
	}

sub	update
	{
	my $self	= shift;
	return $self->save(@_);
	}

#------------------------------------------------------------------------------
# raw file import

sub	raw_import
	{
	my $self	= shift;
	if ($self->{'archive'})
		{
		my $target	= shift;
		unless ($target)
			{
			warn	"[" . __PACKAGE__ . "::raw_import] "	.
				"No target member for import\n";
			return undef;
			}
		$target =~ s/^#//;
		return $self->{'archive'}->raw_import($target, @_);
		}
	else
		{
		warn	"[" . __PACKAGE__ . "::raw_import] "	.
			"No archive for file import\n";
		return undef;
		}
	}

#------------------------------------------------------------------------------
# raw file export

sub	raw_export
	{
	my $self	= shift;
	if ($self->{'archive'})
		{
		my $source	= shift;
		unless ($source)
			{
			warn	"[" . __PACKAGE__ . "::raw_import] "	.
				"Missing source file name\n";
			return undef;
			}
		$source =~ s/^#//;
		return $self->{'archive'}->raw_export($source, @_);
		}
	else
		{
		warn	"[" . __PACKAGE__ . "::raw_import] "	.
			"No archive for file export\n";
		return undef;
		}
	}

#------------------------------------------------------------------------------
# exports the whole content of the document as an XML string

sub	getXMLContent
	{
	my $self	= shift;

	my $xml	= $self->exportXMLElement($self->getRoot);

	return	$self->{'begin'}	. "\n" .
		$xml			. "\n" .
		$self->{'end'};
	}

sub	getContent
	{
	my $self	= shift;
	return $self->getXMLContent;
	}

#------------------------------------------------------------------------------
# brute force tree reorganization

sub	reorganize
	{
	my $self	= shift;
	my $xml		= $self->exportXMLElement($self->getRoot);
	$self->{'xpath'} = $self->{'parser'}->parse($xml);
	}

#------------------------------------------------------------------------------
# returns the root of the XML document

sub	getRoot
	{
	my $self	= shift;

	return $self->getElement('/', 0);
	}

#------------------------------------------------------------------------------
# returns the root element of the XML document

sub	getRootElement
	{
	my $self	= shift;
	return $self->getElement('/' . $self->{'element'}, 0);
	}
		
#------------------------------------------------------------------------------
# returns the content class (text, spreadsheet, presentation, drawing)

sub	contentClass
	{
	my $self	= shift;
	my $class	= shift;

	my $element = $self->getRootElement;
	$self->setAttribute($element, 'office:class', $class) if $class;
	return $self->getAttribute($element, 'office:class');
	}

#------------------------------------------------------------------------------
# member type checks

sub	isContent
	{
	my $self	= shift;
	return ($self->getRootName() eq $XMLNAMES{'content'}) ? 1 : undef;
	}

sub	isMeta
	{
	my $self	= shift;
	return ($self->getRootName() eq $XMLNAMES{'meta'}) ? 1 : undef;
	}
	
sub	isStyles
	{
	my $self	= shift;
	return ($self->getRootName() eq $XMLNAMES{'styles'}) ? 1 : undef;
	}
	
sub	isSettings
	{
	my $self	= shift;
	return ($self->getRootName() eq $XMLNAMES{'settings'}) ? 1 : undef;
	}

#------------------------------------------------------------------------------
# returns the document body element (if defined)

sub	getBody
	{
	my $self	= shift;

	return 
		(
		$self->getElement($self->{'body_path'}, 0)
			||
		$self->getElement($self->{'master_style_path'}, 0)
		);
	}

#------------------------------------------------------------------------------
# makes the current OODoc::XPath object share the same content as another one

sub	cloneContent
	{
	$self	= shift;
	$source	= shift;

	unless ($source && $source->{'xpath'})
		{
		warn "[" . __PACKAGE__ . "]::cloneContent - No valid source\n";
		return undef;
		}
		
	$self->{'xpath'}	= $source->{'xpath'};
	$self->{'begin'}	= $source->{'begin'};
	$self->{'xml'}		= $source->{'xml'};
	$self->{'end'}		= $source->{'end'};

	return $self->getRoot;
	}

#------------------------------------------------------------------------------
# exports an individual element as an XML string

sub	exportXMLElement
	{
	my $self	= shift;
	my $path	= shift;
	my $element	=
		(ref $path) ? $path : $self->getElement($path, @_);

	my $text	= $element->toString;
	return $text;
	}

#------------------------------------------------------------------------------
# exports the document body (if defined) as an XML string

sub	exportXMLBody
	{
	my $self	= shift;

	return	$self->exportXMLElement($self->getBody, @_);
	}

#------------------------------------------------------------------------------
# gets the reference of an XML element identified by path & position
# for subsequent processing

sub	getElement
	{
	my $self	= shift;
	my $path	= shift;
	return undef	unless $path;
	if (ref $path)
		{
		return	$path->isElementNode ? $path : undef;
		}
	my $pos		= shift;
	if (defined $pos && (($pos =~ /^\d*$/) || ($pos =~ /^[\d+-]\d+$/)))
		{
		my $context	= shift;
		$context	= $self->{'xpath'} unless ref $context;
		my $node	= ($context->find($path)->get_nodelist)[$pos];

		return	$node && $node->isElementNode ? $node : undef;
		}
	else
		{
		warn	"[" . __PACKAGE__ . "::getElement] "	.
			"Missing or invalid position\n";
		return undef;
		}
	}

#------------------------------------------------------------------------------
# get the list of children (or the first child unless wantarray) matching
# a given element name and belonging to a given element

sub	selectChildElementsByName
	{
	my $self	= shift;
	my $path	= shift;
	my $element	= ref $path ? $path : $self->getElement($path, shift);
	return undef	unless $element;
	my @list	= $element->getChildNodes;
	my $filter	= shift;

	if ($filter && ($filter ne ".*"))
		{
		my @selection = ();
		while (@list)
			{
			my $node	= shift @list;
			my $n		= $node->getName;
			push @selection, $node	if ($n && ($n =~ /$filter/));
			}
		@list	= @selection;
		}
		
	return undef	unless @list;
	return wantarray ? @list : $list[0];
	}

#------------------------------------------------------------------------------
# get the first child belonging to a given element and matching a given name

sub	selectChildElementByName
	{
	my $self	= shift;
	my $path	= shift;
	my $element	= ref $path ? $path : $self->getElement($path, shift);
	return undef			unless $element;
	my $filter	= shift;
	return $element->getFirstChild	unless $filter;
	my @list	= $element->getChildNodes;

	while (@list)
		{
		my $node	= shift @list;
		my $n		= $node->getName;
		return $node	if ($n && ($n =~ /$filter/));
		}

	return undef;
	}

#------------------------------------------------------------------------------
# get the first child belonging to a given element with an exact given name

sub	getChildElementByName
	{
	my $self	= shift;
	my $path	= shift;
	my $element	= ref $path ? $path : $self->getElement($path, shift);
	return undef			unless $element;
	my $filter	= shift;
	return undef unless $filter;
	my @list	= $element->getChildNodes;
	while (@list)
		{
		my $node	= shift @list;
		my $n		= $node->getName;
		return $node	if ($n && ($n eq $filter));
		}
	return undef;
	}

#------------------------------------------------------------------------------
# replaces any previous content of an existing element by a given text

sub	setText
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $text	= shift;

	return undef	unless defined $text;
	
	my $element 	= $self->getElement($path, $pos);
	return undef	unless $element;

	$element->removeChildNodes;
	my @lines	= split "\n", $text;
	while (@lines)
		{
		my $line	= shift @lines;
		my @columns	= split "\t", $line;
		while (@columns)
			{
			my $column	=
				OpenOffice::OODoc::XPath::encode_text
					(shift @columns);
			$element->appendChild
				(XML::XPath::Node::Text->new($column));
			$self->appendElement($element, 'text:tab-stop')
					if (@columns);
			}
		$self->appendElement($element, 'text:line-break')
				if (@lines);
		}

	return $text;
	}

#------------------------------------------------------------------------------
# extends the text of an existing element

sub	extendText
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $text	= shift;

	return undef	unless defined $text;
	
	my $element 	= $self->getElement($path, $pos);
	return undef	unless $element;

	my @lines	= split "\n", $text;
	while (@lines)
		{
		my $line	= shift @lines;
		my @columns	= split "\t", $line;
		while (@columns)
			{
			my $column	=
				OpenOffice::OODoc::XPath::encode_text
					(shift @columns);
			$element->appendChild
				(XML::XPath::Node::Text->new($column));
			$self->appendElement($element, 'text:tab-stop')
					if (@columns);
			}
		$self->appendElement($element, 'text:line-break')
				if (@lines);
		}

	return $text;
	}

#------------------------------------------------------------------------------
# creates a new encoded text node

sub	createTextNode
	{
	my $self	= shift;
	my $text	= shift		or return undef;
	my $content	= OpenOffice::OODoc::XPath::encode_text($text);
	return XML::XPath::Node::Text->new($content);
	}

#------------------------------------------------------------------------------
# replaces substring in an element 

sub	replaceText
	{
	my $self	= shift;
	my $path	= shift;
	my $element	= (ref $path) ?
				$path	:
				$self->getElement($path, shift);
	
	return	$element ?
		$element->_search_content(@_)	:
		undef;
	}

#------------------------------------------------------------------------------
# gets text in element by path (sub-element texts are concatenated)

sub	getText	
	{
	my $self	= shift;
	my $element	= $self->getElement(@_);
	return undef	unless ($element && $element->isElementNode);
	my $text	= '';
	
	my $name	= $element->getName;
	if	($name eq 'text:tab-stop')	{ return "\t"; }
	if	($name eq 'text:line-break')	{ return "\n"; }
	foreach my $node ($element->getChildNodes)
		{
		if ($node->isElementNode)
			{ $text .= ($self->getText($node) || ''); }
		else
			{
			my $t = ($node->getValue() || '');
			$text .= OpenOffice::OODoc::XPath::decode_text($t);
			}
		}
	return $text;
	}

#------------------------------------------------------------------------------
# returns the children of a given element

sub	getElementList
	{
	my $self	= shift;
	my $path	= shift;
	
	$path		= "/"	unless $path;

	return $self->{'xpath'}->findnodes($path);
	}

#------------------------------------------------------------------------------
# brute XPath nodelist selection; allows any XML::XPath expression

sub	selectNodesByXPath
	{
	my $self	= shift;
	my ($p1, $p2)	= @_;
	my $path	= undef;
	my $context	= undef;
	if (ref $p1)	{ $context = $p1; $path = $p2; }
	else		{ $path = $p1; $context = $p2; }

	$context = $self->{'xpath'} unless ref $context;
	return $context->find($path, @_)->get_nodelist;
	}

#------------------------------------------------------------------------------
# brute XPath single node selection; allows any XML::XPath expression

sub	selectNodeByXPath
	{
	my $self	= shift;
	return ($self->selectNodesByXPath(@_))[0];
	}

sub	getNodeByXPath
	{
	my $self	= shift;
	return ($self->selectNodesByXPath(@_))[0];
	}

#------------------------------------------------------------------------------
# brute XPath value extraction; allows any XML::XPath expression

sub	getXPathValue
	{
	my $self	= shift;
	my ($p1, $p2)	= @_;
	my $path	= undef;
	my $context	= undef;
	if (ref $p1)	{ $context = $p1; $path = $p2; }
	else		{ $path = $p1; $context = $p2; }
	 
	if (ref $context)
		{
		$path =~ s/^\/*//;
		return OpenOffice::OODoc::XPath::decode_text
				($context->findvalue($path, @_));
		}
	else
		{
		return OpenOffice::OODoc::XPath::decode_text
				($self->{'xpath'}->findvalue($path, @_));
		}
	}

#------------------------------------------------------------------------------
# create or update an xpath

sub	makeXPath
	{
	my $self	= shift;
	my $path	= shift;
	my $root	= undef;
	if (ref $path)
		{
		$root	= $path;
		$path	= shift;
		}
	else
		{
		$root	= $self->getRoot;
		}
	$path =~ s/^[\/ ]*//; $path =~ s/[\/ ]*$//;
	my @list	= split '/', $path;
	my $posnode	= $root;
	while (@list)
		{
		my $item	= shift @list;
		while (($item =~ /\[.*/) && !($item =~ /\[.*\]/))
			{
			my $cont = shift @list or last;
			$item .= ('/' . $cont);
			}
		next unless $item;
		my $node	= undef;
		my $name	= undef;
		my $param	= undef;
		$item =~ s/\[(.*)\] *//;
		$param = $1;
		$name = $item; $name =~ s/^ *//; $name =~ s/ *$//;
		my %attributes = ();
		my $text = undef;
		my $indice = undef;
		if ($param)
			{
			my @attrlist = [];
			$indice = undef;
			$param =~ s/^ *//; $param =~ s/ *$//;
			$param =~ s/^@//;
			@attrlist = split /@/, $param;
			foreach my $a (@attrlist)
				{
				next unless $a;
				$a =~ s/^ *//;
				my $tmp = $a;
				$tmp =~ s/ *$//;
				if ($tmp =~ /^\d*$/)
					{
					$indice = $tmp;
					next;
					}
				if ($a =~ s/^\"(.*)\".*/$1/)
					{
					$text = $1; next;
					}
				if ($a =~ /^=/)
					{
					$a	=~ s/^=//;
					$a	=~ '^"(.*)"$';
					$text	= $1 ? $1 : $a;
					next;
					}
				$a =~ s/^@//;
				my ($attname, $attvalue) = split '=', $a;
				next unless $attname;
				if ($attvalue)
					{
					$attvalue =~ '"(.*)"';
					$attvalue = $1 if $1;
					}
				$attname =~ s/^ *//; $attname =~ s/ *$//;
				$attributes{$attname} = $attvalue;
				}
			}
		if (defined $indice)
			{
			$node = $self->getNodeByXPath
					($posnode, "$name\[$indice\]");
			}
		else
			{
			$node	=
				$self->getChildElementByName($posnode, $name);
			}
		if ($node)
			{
			$self->setAttributes($node, %attributes);
			$self->setText($node, $text)	if (defined $text);
			}
		else
			{
			$node = $self->appendElement
					(
					$posnode, $name,
					text		=> $text,
					attributes	=> {%attributes}
					);
			}
		if ($node)	{ $posnode = $node;	}
		else		{ return undef;		}
		}
	return $posnode;
	}

#------------------------------------------------------------------------------
# selects element by path and attribute

sub	selectElementByAttribute
	{
	my $self	= shift;
	my $path	= shift;
	my $key		= shift;
	my $value	= shift || '';

	my @candidates	=  $self->getElementList($path);
	return @candidates	unless $key;

	for (@candidates)
		{
		if ($_->isElementNode)
			{
			my $v = $self->getAttribute($_, $key);
			return $_ if (defined $v && ($v =~ /$value/));
			}
		}
	return undef;
	}

#------------------------------------------------------------------------------
# selects list of elements by path and attribute

sub	selectElementsByAttribute
	{
	my $self	= shift;
	my $path	= shift;
	my $key		= shift;
	my $value	= shift || '';

	my @candidates	=  $self->getElementList($path);
	return @candidates	unless $key;

	my @selection	= ();
	for (@candidates)
		{
		if ($_->isElementNode)
			{
			my $v	= $self->getAttribute($_, $key);
			push @selection, $_	if ($v && ($v =~ /$value/));
			}
		}

	return wantarray ? @selection : $selection[0];
	}

#------------------------------------------------------------------------------
# get a list of elements matching a given path and an optional content pattern

sub	findElementList
	{
	my $self	= shift;
	my $path	= shift;
	my $pattern	= shift;
	my $replace	= shift;

	return undef unless $path;

	my @result	= ();

	foreach my $n ($self->{'xpath'}->findnodes($path))
		{
		push @result,
		    [ $self->findDescendants($n, $pattern, $replace, @_) ];
		}

	return @result;
	}

#------------------------------------------------------------------------------
# get a list of elements matching a given path and an optional content pattern
# without replacement operation, and from an optional context node

sub	selectElements
	{
	my $self	= shift;
	my $path	= shift;
	my $context	= $self->{'xpath'};
	if (ref $path)
		{
		$context	= $path;
		$path		= shift;
		}
	my $filter	= shift;

	my @candidates	= $self->selectNodesByXPath($context, $path);
	return @candidates	unless $filter;

	my @result	= ();
	while (@candidates)
		{
		my $node = shift @candidates;
		push @result, $node
			if $node->_search_content($filter, @_, $node);
		}
	return @result;
	}

#------------------------------------------------------------------------------
# get the 1st element matching a given path and on optional content pattern

sub	selectElement
	{
	my $self	= shift;
	my $path	= shift;
	my $context	= $self->{'xpath'};
	if (ref $path)
		{
		$context	= $path;
		$path		= shift;
		}
	return undef	unless $path;
	my $filter	= shift;

	my @candidates	= $self->selectNodesByXPath($context, $path);
	return $candidates[0]	unless $filter;

	while (@candidates)
		{
		my $node = shift @candidates;
		return $node if $node->_search_content($filter, @_, $node);
		}
	return undef;
	}

#------------------------------------------------------------------------------
# gets the descendants of a given node, with optional in fly search/replacement

sub	findDescendants
	{
	my $self	= shift;
	my $node	= shift;
	my $pattern	= shift;
	my $replace	= shift;
	
	my @result		= ();

	my $n	= $self->selectNodeByContent($node, $pattern, $replace, @_);
	push @result, $n	if $n;
	foreach my $m ($node->getChildNodes)
		{
		push @result,
		    [ $self->findDescendants($m, $pattern, $replace, @_) ];
		}

	return @result;
	}

#------------------------------------------------------------------------------
# search & replace text in an individual node

sub	selectNodeByContent
	{
	my $self	= shift;
	my $node	= shift;
	my $pattern	= shift;
	my $replace	= shift;

	return $node	unless $pattern;
	my $l	= $node->getNodeValue;

	return undef	unless $l;

	unless (defined $replace)
		{
		return ($l =~ /$pattern/) ? $node : undef;
		}
	else
		{
		if (ref $replace)
			{
			unless
			  ($l =~ s/($pattern)/&$replace(@_, $node, $1)/eg)
				{
				return undef;
				}
			}
		else
			{
			unless ($l =~ s/$pattern/$replace/g)
				{
				return undef;
				}
			}
		$node->setNodeValue($l);
		return $node;
		}
	}

#------------------------------------------------------------------------------
# gets the text content of a nodelist

sub	getTextList
	{
	my $self	= shift;
	my $path	= shift;
	my $pattern	= shift;

	return undef unless $path;

	my @nodelist = $self->{'xpath'}->findnodes($path);
	my @text = ();

	foreach my $n (@nodelist)
		{
		my $l	= OpenOffice::OODoc::XPath::decode_text
							($n->string_value);
		push @text, $l if ((! defined $pattern) || ($l =~ /$pattern/));
		}

	return wantarray ? @text : join "\n", @text;
	}

#------------------------------------------------------------------------------
# gets the attributes of an element in the key => value form

sub	getAttributes
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;

	return undef	unless $path;
	$pos	= 0	unless $pos;

	my $node	= $self->getElement($path, $pos);
	return undef	unless $path;

	my %attributes	= ();
	foreach my $a ($node->getAttributeNodes)
		{
		my $name		= $a->getName;
		$attributes{$name}	=
			OpenOffice::OODoc::XPath::decode_text($a->getValue);
		}

	return %attributes;
	}

#------------------------------------------------------------------------------
# gets the value of an attribute by path + name

sub	getAttribute
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $name	= shift;

	return undef	unless $path;
	$pos	= 0	unless $pos;

	my $node	= $self->getElement($path, $pos);
	return	OpenOffice::OODoc::XPath::decode_text
					($node->getAttribute($name));
	}

#------------------------------------------------------------------------------
# set/replace a list of attributes in an element

sub	setAttributes
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my %attr	= @_;

	my $node	= $self->getElement($path, $pos);

	return undef	unless $node;

	foreach my $k (keys %attr)
		{
		if (defined $attr{$k})
		    {
		    $node->setAttribute
		    		(
				$k,
				OpenOffice::OODoc::XPath::encode_text
						($attr{$k})
				);
		    }
		elsif (my $a = $node->getAttributeNode($k))
		    {
		    $node->removeAttribute($a);
		    }
		}

	return %attr;
	}

#------------------------------------------------------------------------------
# set/replace a single attribute in an element

sub	setAttribute
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $node	= $self->getElement($path, $pos) or return undef;
	my $attribute	= shift or return undef;
	my $value	= shift;

	if (defined $value && ($value gt ' '))
		{
		$node->setAttribute
			(
			$attribute,
			OpenOffice::OODoc::XPath::encode_text($value)
			);
		}
	elsif (my $a = $node->getAttributeNode($attribute))
		{
		$node->removeAttribute($a);
		}
	
	return $value; 
	}

#------------------------------------------------------------------------------
# removes an attribute in element

sub	removeAttribute
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $name	= shift;

	my $node	= $self->getElement($path, $pos);

	return undef	unless $node;

	foreach my $a ($node->getAttributeNodes)
		{
		if ($a->getName eq $name)
			{
			$node->removeAttribute($a);
			}
		}

	return 1;
	}

#------------------------------------------------------------------------------
# replicates an existing element, provided as an XPath ref or an XML string

sub	replicateElement
	{
	my $self	= shift;
	my $proto	= shift;
	my $position	= shift;
	my %options	= @_;

	unless ($proto && ref $proto && $proto->isElementNode)
		{
		warn "[" . __PACKAGE__ . "::replicateElement] No prototype\n";
		return undef;
		}

	$position	= 'end'	unless $position;

	my $element		= undef;
	my $name		= $proto->getName;
	%{$options{'attribute'}} = $self->getAttributes($proto);

	if	(ref $position)
		{
		if (! $options{'position'})
			{
			$element = $self->appendElement
						($position, $name, %options);
			}
		else
			{
			$element = $self->insertElement
						($position, $name, %options);
			}
		}
	elsif	($position eq 'end')
		{
		$element = $self->appendElement
					($self->getRoot, $name, %options);
		}
	elsif	($position eq 'body')
		{
		$element = $self->appendElement
					($self->getBody, $name, %options);
		}

	foreach my $node ($proto->getChildNodes)
		{
		if	($node->isElementNode)
			{
			$options{'position'} = undef;
			$self->replicateElement($node, $element, %options);
			}
		elsif	($node->isTextNode)
			{
			my $text_node = XML::XPath::Node::Text->new
							($node->getValue);
			$element->appendChild($text_node);
			}
		}
	return $element;
	}

#------------------------------------------------------------------------------
# create an element, just with a mandatory name and an optional text
# the name can have the namespace:name form
# if the $name argument is a '<.*>' string, it's processed as XML and
# the new element is completely generated from it

sub	createElement
	{
	my $self	= shift;
	my $name	= shift;
	my $text	= shift;
	my $element	= undef;

	unless ($name)
		{
		warn "[" . __PACKAGE__ . "::createElement] No name or XML\n";
		return undef;
		}
	$name =~ s/^\s+//;
	$name =~ s/\s+$//;
	if ($name =~ /^<.*>$/)
		{
		$element =  $self->{'parser'}->parse($name);
		return ($element && $element->isElementNode)	?
			$element : undef;
		}
	else
		{
		$name		=~ /(.*):(.*)/;
		my $prefix	=  $1;
		$element = XML::XPath::Node::Element->new($name, $prefix);
		}

	$self->setText($element, $text)		if ($text);

	return $element;
	}
	
#------------------------------------------------------------------------------
# replaces an element by another one
# the new element is inserted before the old one,
# then the old element is removed.
# the new element can be inserted by copy (default) or by reference
# return = new element if success, undef if failure

sub	replaceElement
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $new_element	= shift;
	my %options	=
			(
			mode		=> 'copy',
			@_
			);
	unless ($new_element)
		{
		warn	"[" . __PACKAGE__ . "::replaceElement] " .
			"Missing new element\n";
		return undef;
		}
	unless (ref $new_element)
		{
		$new_element = $self->createElement($new_element);
		$options{'mode'} = 'reference';
		}
	unless ($new_element && $new_element->isElementNode)
		{
		warn	"[" . __PACKAGE__ . "::replaceElement] " .
			"No valid replacement\n";
		return undef;
		}

	my $result	= undef;

	my $old_element	= $self->getElement($path, $pos);
	unless ($old_element)
		{
		warn	"[" . __PACKAGE__ . "::replaceElement] " .
			"Non existing element to be replaced\n";
		return undef;
		}
	if	(! $options{'mode'} || $options{'mode'} eq 'copy')
		{
		$result = $self->replicateElement
					(
					$new_element,
					$old_element,
					position	=> 'before'
					);
		}
	elsif	($options{'mode'} && $options{'mode'} eq 'reference')
		{
		$result = $self->insertElement
					(
					$old_element,
					$new_element,
					position	=> 'before'
					);
		}
	else
		{
		warn	"[" . __PACKAGE__ . "::replaceElement] " .
			"Unknown option\n";
		return undef;
		}
	if	($result && $result->isElementNode)
		{
		$self->removeElement($old_element);
		return $result;
		}
	return undef;
	}

#------------------------------------------------------------------------------
# adds a new or existing child element

sub	appendElement
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $name	= shift;
	my %opt		= @_;

	$opt{'attribute'} = $opt{'attributes'} unless ($opt{'attribute'});

	return undef	unless $name;
	my $element	= undef;
	
	unless (ref $name)
		{
		$element	= $self->createElement($name, $opt{'text'});
		}
	else
		{
		$element	= $name;
		$self->setText($element, $opt{'text'})	if $opt{'text'};
		}
	return undef	unless $element;
	my $parent	= $self->getElement($path, $pos);
	unless ($parent)
		{
		warn	"[" . __PACKAGE__ .
			"::appendElement] Position not found\n";
		return undef;
		}
	$parent->appendChild($element);
	$self->setAttributes($element, %{$opt{'attribute'}});
	return $element;
	}

#------------------------------------------------------------------------------
# inserts a new element before or after a given node
# as appendElement, but the new element is a 'brother' (and not a child) of
# the first given element

sub	insertElement
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $name	= shift;
	my %opt		= @_;

	return undef	unless $name;
	my $element	= undef;
	unless (ref $name)
		{
		$element	= $self->createElement($name, $opt{'text'});
		}
	else
		{
		$element	= $name;
		$self->setText($element, $opt{'text'})	if $opt{'text'};
		}
	return undef	unless $element;
	
	my $posnode	= $self->getElement($path, $pos);
	unless ($posnode)
		{
		warn "[" . __PACKAGE__ . "::insertElement] Unknown position\n";
		return undef;
		}
	my $parent	= $posnode->getParentNode;
	unless ($parent)
		{
		warn "[" . __PACKAGE__ . "::insertElement] Root position\n";
		return undef;
		}

	if (($opt{'position'}) && ($opt{'position'} eq 'after'))
		{
		$parent->insertAfter($element, $posnode);
		}
	else
		{
		$parent->insertBefore($element, $posnode);
		}

	$self->setAttributes($element, %{$opt{'attribute'}});

	return $element;
	}

#------------------------------------------------------------------------------
# removes the given element & children

sub	removeElement
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;

	my $e	= $self->getElement($path, $pos);
	return undef	unless $e;
	my $p	= $e->getParentNode;

	unless ($p)
		{
		warn "[" . __PACKAGE__ . "::removeElement] Root node\n";
		return undef;
		}
	$p->removeChild($e);

	return 1;
	}

#------------------------------------------------------------------------------
1;
