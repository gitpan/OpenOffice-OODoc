#-----------------------------------------------------------------------------
#
#	$Id : XPath.pm 1.200 2005-02-07 JMG$
#
#	Initial developer: Jean-Marie Gouarne
#	Copyright 2005 by Genicorp, S.A. (www.genicorp.com)
#	Licensing conditions:
#		- Licence Publique Generale Genicorp v1.0
#		- GNU Lesser General Public License v2.1
#	Contact: oodoc@genicorp.com
#
#-----------------------------------------------------------------------------

package	OpenOffice::OODoc::XPath;
use	5.008_000;
our	$VERSION	= 1.200;
use	XML::Twig	3.15;
use	Encode;

#------------------------------------------------------------------------------

our %XMLNAMES	=			# OODoc root element names
	(
	'content'	=> 'office:document-content',
	'styles'	=> 'office:document-styles',
	'meta'		=> 'office:document-meta',
	'manifest'	=> 'manifest:manifest',
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
# basic element creation

sub	OpenOffice::OODoc::XPath::new_element
	{
	my $name	= shift		or return undef;
	return undef if ref $name;

	$name		=~ s/^\s+//;
	$name		=~ s/\s+$//;
	
	if ($name =~ /^<.*>$/)	# create element from XML string
		{
		return XML::Twig::Elt->parse($name, @_);
		}
	else			# create element from name and optional data
		{
		return XML::Twig::Elt->new($name, @_);
		}
	}

#------------------------------------------------------------------------------
# text node creation

sub	OpenOffice::OODoc::XPath::new_text_node
	{
	return OpenOffice::OODoc::XPath::new_element('#PCDATA');
	}

#------------------------------------------------------------------------------
# some wrappers for XML Twig elements
package	XML::Twig::Elt;
#------------------------------------------------------------------------------

sub	setName
	{
	my $node	= shift;
	return $node->set_tag(shift);
	}

sub	getPrefix
	{
	my $node	= shift;
	return $node->ns_prefix;
	}

sub	selectChildElements
	{
	my $node	= shift;
	my $filter	= shift;

	return $node->getChildNodes() unless $filter;

	my @list	= ();
	my $fc = $node->first_child;
	my $name = $fc->name if $fc;
	while ($fc)
		{
		if ($name && ($name =~ /$filter/))
			{
			return $fc unless wantarray;
			push @list, $fc;
			}
		$fc = $fc->next_sibling;
		$name = $fc->name if $fc;;
		}
	return @list;
	}

sub	selectChildElement
	{
	my $node	= shift;
	return scalar $node->selectChildElements(@_);
	}

sub	getParentNode
	{
	my $node	= shift;
	return $node->parent;
	}
	
sub	getFirstChild
	{
	my $node	= shift;
	my $fc = $node->first_child(@_);
	my $name = $fc->name if $fc;
	while ($name && ($name =~ /^#/))
		{
		$fc = $fc->next_sibling(@_);
		$name = $fc->name if $fc;
		}
	return $fc;
	}

sub	getLastChild
	{
	my $node	= shift;
	my $lc = $node->last_child(@_);
	my $name = $lc->name;
	while ($name && ($name =~ /^#/))
		{
		$lc = $lc->prev_sibling(@_);
		$name = $lc->name;
		}
	return $lc;
	}

sub	getDescendantNodes
	{
	my $node	= shift;
	my @children	= $node->getChildNodes(@_);
	my @descendants = ();
	foreach my $child (@children)
		{
		push @descendants, $child, $child->getDescendantNodes(@_);
		}
	return @descendants;
	}

sub	getDescendantTextNodes
	{
	my $node	= shift;
	return $node->descendants('#PCDATA');
	}

sub	appendChild
	{
	my $node	= shift;
	my $child	= shift;
	unless (ref $child)
		{
		$child = OpenOffice::OODoc::XPath::new_element($child, @_);
		}
	return $child->paste_last_child($node);
	}

sub	removeChildNodes
	{
	my $node	= shift;
	return $node->cut_children(@_);
	}

sub	replicateNode
	{
	my $node	= shift;
	my $number	= shift || 1;
	my $position	= shift || 'after';
	my $lastnode	= $node;
	while ($number > 0)
		{
		my $newnode	= $node->copy;
		$newnode->paste($position => $lastnode);
		$last_node	= $newnode;
		$number--;
		}
	return $last_node;
	}

sub	getNodeValue
	{
	my $node	= shift;
	return $node->text;
	}

sub	getValue
	{
	my $node	= shift;
	return $node->text;
	}

sub	setNodeValue
	{
	my $node	= shift;
	return $node->set_text(@_);
	}

sub	appendTextChild
	{
	my $node	= shift;
	my $text	= shift or return undef;
	my $text_node	= XML::Twig::Elt->new('#PCDATA' => $text);
	return $text_node->paste(last_child => $node);
	}

sub	getAttribute
	{
	my $node	= shift;
	return $node->att(@_);
	}

sub	getAttributes
	{
	my $node	= shift;
	return %{$node->atts(@_)};
	}

sub	setAttribute
	{
	my $node	= shift or return undef;
	my $attribute	= shift;
	my $value	= shift;
	if (defined $value)
		{
		return $node->set_att($attribute, $value, @_);
		}
	else
		{
		if ($node->{$attribute})
			{
			return $node->del_att($attribute);
			}
		else
			{
			return undef;
			}
		}
	}

sub	removeAttribute
	{
	my $node	= shift or return undef;
	return $node->del_att(shift);
	}

#------------------------------------------------------------------------------
package	OpenOffice::OODoc::XPath;
#------------------------------------------------------------------------------
# basic conversion between internal & printable encodings (object version)

sub	inputTextConversion
	{
	my $self	= shift;
	my $text	= shift;
	my $local_encoding = $self->{'local_encoding'} or return $text;
	return Encode::decode($local_encoding, $text);
	}

sub	outputTextConversion
	{
	my $self	= shift;
	my $text	= shift;
	my $local_encoding = $self->{'local_encoding'} or return $text;
	return Encode::encode($local_encoding, $text);
	}

sub	localEncoding
	{
	my $self	= shift;
	my $encoding	= shift;
	$self->{'local_encoding'} = $encoding if $encoding;
	return $self->{'local_encoding'} || '';
	}

sub	noLocalEncoding
	{
	my $self	= shift;
	delete $self->{'local_encoding'};
	return 1;
	}

#------------------------------------------------------------------------------
# search/replace text processing routine
# if $replace is a user-provided routine, it's called back with
# the current argument stack, plus the substring found

sub	_find_text
	{
	my $self	= shift;
	my $text	= shift;
	my $pattern	= $self->inputTextConversion(shift);
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
		    my $r = $self->inputTextConversion($replace);
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
# search/replace content in descendant nodes

sub	_search_content
	{
	my $self	= shift;
	my $element	= shift or return undef;
	my $content	= undef;
	
	foreach my $child ($element->children)
		{
		my $text = undef;
		if ($child->isTextNode)
			{
			$text = $self->_find_text($child->text, @_);
			$child->set_text($text) if defined $text;
			}
		else
			{
			my $t = $self->_search_content($child, @_);
			$text .= $t if defined $t;
			}
		$content .= $text if defined $text;
		}
	return $content;
	}

#------------------------------------------------------------------------------
# document class check

sub	isContent
	{
	my $self	= shift;
	return ($self->contentClass()) ? 1 : undef;
	}

sub	isCalcDocument
	{
	my $self	= shift;
	return ($self->contentClass() eq 'spreadsheet') ? 1 : undef;
	}
sub	isImpressDocument
	{
	my $self	= shift;
	return ($self->contentClass() eq 'impress') ? 1 : undef;
	}
sub	isDrawDocument
	{
	my $self	= shift;
	return ($self->contentClass() eq 'drawing') ? 1 : undef;
	}
sub	isWriterDocument
	{
	my $self	= shift;
	return ($self->contentClass() eq 'text') ? 1 : undef;
	}

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
		local_encoding		=>
				$OpenOffice::OODoc::XPath::LOCAL_CHARSET,
		@_
		};

	my $twig = undef;

	if ($self->{'member'} && ! $self->{'element'})
		{
		my $m	= lc $self->{'member'};
		$m	=~ /(^.*)\..*/;
		$m	= $1	if $1;
		$self->{'element'} =
		    $OpenOffice::OODoc::XPath::XMLNAMES{$m};
		}
					# create the XML::Twig
	if ($self->{'readable_XML'} && ($self->{'readable_XML'} eq 'on'))
		{
		$self->{'readable_XML'} = 'indented';
		}
	$self->{'element'} = $OpenOffice::OODoc::XPath::XMLNAMES{'content'}
				unless $self->{'element'};
	if ($self->{'element'})
		{
		$twig = XML::Twig->new
			(
			twig_roots	=>
				{
				$self->{'element'}	=> 1
				},
			pretty_print	=> $self->{'readable_XML'},
			%{$self->{'twig_options'}}
			);
		}
	else
		{
		$twig = XML::Twig->new
			(
			pretty_print	=> $self->{'readable_XML'},
			%{$self->{'twig_options'}}
			);
		}

	if ($self->{'xml'})		# load from XML string
		{
		$self->{'xpath'} = $twig->safe_parse($self->{'xml'});
		}
	elsif ($self->{'archive'})	# load from existig OOFile
		{
		$self->{'member'} = 'content' unless $self->{'member'};
		$self->{'xml'} = $self->{'archive'}->link($self);
		$self->{'xpath'} = $twig->safe_parse($self->{'xml'});
		}
	elsif ($self->{'file'})		# look for a file
		{
		unless	(
			$self->{'flat_xml'}
					||
			(lc $self->{'file'}) =~ /\.xml$/
			)
			{		# create OOFile & extract from it
			require OpenOffice::OODoc::File;

			$self->{'archive'} = OpenOffice::OODoc::File->new
				(
				$self->{'file'},
				create		=> $self->{'create'},
				template_path	=> $self->{'template_path'}
				);
			$self->{'member'} = 'content'
					unless $self->{'member'};
			$self->{'xml'} = $self->{'archive'}->link($self);
			$self->{'xpath'} = $twig->safe_parse($self->{'xml'});
			}
		else
			{		# load from XML flat file
			$self->{'xpath'} =
				$twig->safe_parsefile($self->{'file'});		
			}
		}
	else
		{
		warn "[" . __PACKAGE__ . "] No XML content\n";
		return undef;
		}
	delete $self->{'xml'};
	unless ($self->{'xpath'})
		{
		warn "[" . __PACKAGE__ . "] No well formed content\n";
		return undef;
		}
	return bless $self, $class;
	}

#------------------------------------------------------------------------------

sub	DESTROY
	{
	my $self	= shift;

	delete $self->{'xpath'};
	delete $self->{'archive'};
	$self = {};
	}

sub	dispose
	{
	my $self	= shift;
	$self->DESTROY(@_);
	}

#------------------------------------------------------------------------------
# get a reference to the embedded XML parser for share

sub	getXMLParser
	{
	warn	"[" . __PACKAGE__ . "::getXMLParser] No longer implemented\n";
	return undef;
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
	return $self->getRootElement()->sprint;
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
	warn "[" . __PACKAGE__ . "::reorganize] No longer implemented\n";
	return undef;
	}

#------------------------------------------------------------------------------
# returns the root of the XML document

sub	getRoot
	{
	my $self	= shift;
	return $self->{'xpath'}->root;
	}

#------------------------------------------------------------------------------
# returns the root element of the XML document

sub	getRootElement
	{
	my $self	= shift;
	my $root	= $self->{'xpath'}->root;
	my $rootname	= $root->name() || '';
	return ($rootname eq $self->{'element'})	?
			$root				:
			$root->first_child($self->{'element'});
	}
		
#------------------------------------------------------------------------------
# returns the content class (text, spreadsheet, presentation, drawing)

sub	contentClass
	{
	my $self	= shift;
	my $class	= shift;

	my $element = $self->getRootElement;
	return undef unless $element;
	$self->setAttribute($element, 'office:class', $class) if $class;
	return $self->getAttribute($element, 'office:class') || '';
	}

#------------------------------------------------------------------------------
# element name check

sub	getRootName
	{
	my $self	= shift;
	return $self->getRootElement->name;
	}

#------------------------------------------------------------------------------
# member type checks

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
	my $body	= $self->getRootElement->selectChildElement
				(
				'office:(body|meta|master-styles|settings)'
				);
	return $body;
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

	my $text	= $element->sprint(@_);
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
		$context	= $self->getRootElement unless ref $context;
		my $node	= ($context->findnodes($path))[$pos];

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

	return $element->selectChildElements(@_);
	}

#------------------------------------------------------------------------------
# get the first child belonging to a given element and matching a given name

sub	selectChildElementByName
	{
	my $self	= shift;
	my $path	= shift;
	my $element	= ref $path ? $path : $self->getElement($path, shift);
	return undef			unless $element;
	return scalar $element->selectChildElements(@_);
	}

#------------------------------------------------------------------------------
# get the first child belonging to a given element with an exact given name

sub	getChildElementByName
	{
	my $self	= shift;
	return $self->selectChildElementByName(@_);
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
	
	my $element 	= $self->OpenOffice::OODoc::XPath::getElement
					($path, $pos);
	return undef	unless $element;

	$element->set_text("");
	my @lines	= split "\n", $text;
	while (@lines)
		{
		my $line	= shift @lines;
		my @columns	= split "\t", $line;
		while (@columns)
			{
			my $column	= 
				$self->inputTextConversion(shift @columns);
			$element->appendTextChild($column);
			$element->appendChild('text:tab-stop') if (@columns);
			}
		$element->appendChild('text:line-break') if (@lines);
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
				$self->inputTextConversion(shift @columns);
			$element->appendTextChild($column);
			$element->appendChild('text:tab-stop') if (@columns);
			}
		$element->appendChild('text:line-break') if (@lines);
		}

	return $text;
	}

#------------------------------------------------------------------------------
# creates a new encoded text node

sub	createTextNode
	{
	my $self	= shift;
	my $text	= shift		or return undef;
	my $content	= $self->inputTextConversion($text);
	return XML::Twig::Elt->new('#PCDATA' => $text);
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

	return $self->_search_content($element, @_);
	}

#------------------------------------------------------------------------------
# gets text in element by path (sub-element texts are concatenated)

sub	getText	
	{
	my $self	= shift;
	my $element	= $self->OpenOffice::OODoc::XPath::getElement(@_);
	return undef	unless ($element && $element->isElementNode);
	my $text	= '';
	
	my $name	= $element->getName;
	if	($name eq 'text:tab-stop')	{ return "\t"; }
	if	($name eq 'text:line-break')	{ return "\n"; }
	foreach my $node ($element->getChildNodes)
		{
		if ($node->isElementNode)
			{
			$text .=
			    (
			    $self->getText($node)
			    		||
			    ''
			    );
			}
		else
			{
			my $t = ($node->getValue() || '');
			$text .= $self->outputTextConversion($t);
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
	return $context->get_xpath($path);
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
		}
	else
		{
		$context = $self->{'xpath'};
		}
	return $self->outputTextConversion($context->findvalue($path, @_));
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
			if $self->_search_content($node, $filter, @_, $node);
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
		return $node
			if $self->_search_content($node, $filter, @_, $node);
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
		my $l	= $self->outputTextConversion($n->string_value);
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
	my %atts	= %{$node->atts(@_)};
	foreach my $a (keys %atts)
		{
		$attributes{$a}	= $self->outputTextConversion($atts{$a});
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
	return	$self->outputTextConversion($node->att($name));
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
		    $node->set_att
		    		(
				$k,
				$self->inputTextConversion($attr{$k})
				);
		    }
		else
		    {
		    $node->del_att($a);
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
		$node->set_att
			(
			$attribute,
			$self->inputTextConversion($value)
			);
		}
	else
		{
		$node->del_att($a);
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
	return $node->del_att($a);
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

	my $element		= $proto->copy;
	$element->set_atts($options => 'attribute');

	if	(ref $position)
		{
		if (! $options{'position'})
			{
			$element->paste_last_child($position);
			}
		elsif ($options{'position'} eq 'before')
			{
			$element->paste_before($position);
			}
		elsif ($options{'position'} eq 'after')
			{
			$element->paste_after($position);
			}
		elsif ($options{'position'} ne 'free')
			{
			warn	"[" . __PACKAGE__ . "::replicateElement] " .
				"No valid attachment option\n";
			}
		}
	elsif	($position eq 'end')
		{
		$element->paste_last_child($self->{'xpath'}->root);
		}
	elsif	($position eq 'body')
		{
		$element->paste_last_child($self->getBody);
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

	my $element = OpenOffice::OODoc::XPath::new_element($name, @_);
	unless ($element)
		{
		warn	"[" . __PACKAGE__ . "::createElement] "	.
			"Element creation failure\n";
		return undef;
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
		$result = $new_element->copy;
		$result->replace($old_element);
		return $result;
		}
	elsif	($options{'mode'} && $options{'mode'} eq 'reference')
		{
		$result = $self->insertElement
					(
					$old_element,
					$new_element,
					position	=> 'before'
					);
		$self->removeElement($old_element);
		return $result;
		}
	else
		{
		warn	"[" . __PACKAGE__ . "::replaceElement] " .
			"Unknown option\n";
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
	$element->paste_last_child($parent);
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

	if (($opt{'position'}) && ($opt{'position'} eq 'after'))
		{
		$element->paste_after($posnode);
		}
	else
		{
		$element->paste_before($posnode);
		}

	$self->setAttributes($element, %{$opt{'attribute'}});

	return $element;
	}

#------------------------------------------------------------------------------
# removes the given element & children

sub	removeElement
	{
	my $self	= shift;

	my $e	= $self->getElement(@_);
	return undef	unless $e;
	return $e->delete;
	}

#------------------------------------------------------------------------------
# cuts the given element & children (to be pasted elsewhere)

sub	cutElement
	{
	my $self	= shift;

	my $e	= $self->getElement(@_);
	return undef	unless $e;
	$e->cut;

	return $e;
	}

#------------------------------------------------------------------------------
1;
