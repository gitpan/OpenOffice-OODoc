#-----------------------------------------------------------------------------
#
#	$Id : Text.pm 1.203 2005-02-18 JMG$
#
#	Initial developer: Jean-Marie Gouarne
#	Copyright 2005 by Genicorp, S.A. (www.genicorp.com)
#	Licensing conditions:
#		- Licence Publique Generale Genicorp v1.0
#		- GNU Lesser General Public License v2.1
#	Contact: oodoc@genicorp.com
#
#-----------------------------------------------------------------------------

package OpenOffice::OODoc::Text;
use	5.006_001;
use	OpenOffice::OODoc::XPath	1.202;
our	@ISA		= qw ( OpenOffice::OODoc::XPath );
our	$VERSION	= 1.203;

#-----------------------------------------------------------------------------
# default text style attributes

our	%DEFAULT_TEXT_STYLE	=
	(
	references	=>
		{
		'style:name'			=> undef,
		'style:family'			=> 'paragraph',
		'style:parent-style-name'	=> 'Standard',
		'style:next-style-name'		=> 'Standard',
		'style:class'			=> 'text'
		},
	properties	=>
		{
		}
	);

#-----------------------------------------------------------------------------
# default delimiters for flat text export

our	%DEFAULT_DELIMITERS	=
	(
	'text:footnote-citation'	=>
		{
		begin	=>	'[',
		end	=>	']'
		},
	'text:footnote-body'		=>
		{
		begin	=>	'{NOTE: ',
		end	=>	'}'
		},
	'text:span'			=>
		{
		begin	=>	'<<',
		end	=>	'>>'
		},
	'text:list-item'		=>
		{
		begin	=>	'- ',
		end	=>	''
		},
	);

#-----------------------------------------------------------------------------

our $ROW_REPEAT_ATTRIBUTE       = 'table:number-rows-repeated';
our $COL_REPEAT_ATTRIBUTE       = 'table:number-columns-repeated';

#-----------------------------------------------------------------------------
# constructor

sub	new
	{
	my $caller	= shift;
	my $class	= ref($caller) || $caller;
	my %options	=
		(
		member		=> 'content',	# default XML member
		paragraph_style	=> 'Standard',	# default paragraph style
		header_style	=> 'Heading 1',	# default header style
		use_delimiters	=> 'on',	# use text output delimiters
		field_separator	=> ';',		# table cell separator
		line_separator	=> "\n",	# text line break
		max_rows	=> 32,		# last row in spreadsheets
		max_cols	=> 26,		# last col in spreadsheets
		delimiters	=>
			{ %OpenOffice::OODoc::Text::DEFAULT_DELIMITERS },
		@_
		);

	my $object	= $class->SUPER::new(%options);

	if ($object)
		{
		bless $object, $class;
		}
	return $object;
	}

#-----------------------------------------------------------------------------
# text element type detection (add-in for XML::Twig::Elt)

sub	XML::Twig::Elt::isOrderedList
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:ordered-list')) ?  1 : undef;
	}

sub	XML::Twig::Elt::isUnorderedList
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:unordered-list')) ?  1 : undef;
	}

sub	XML::Twig::Elt::isItemList
	{
	my $element	= shift;
	my $name	= $element->getName;
	return (
		$name &&
			($name eq 'text:ordered-list')
				||
			($name eq 'text:unordered-list')
		)	?
		1 : undef;
	}

sub	XML::Twig::Elt::isListItem
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:list-item')) ?  1 : undef;
	}

sub	XML::Twig::Elt::isParagraph
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:p')) ? 1 : undef;
	}

sub	XML::Twig::Elt::isHeader
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:h')) ? 1 : undef;
	}

sub	XML::Twig::Elt::headerLevel
	{
	my $element	= shift;
	return $element->getAttribute('text:level');
	}

sub	XML::Twig::Elt::isTable
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'table:table')) ? 1 : undef;
	}

sub	XML::Twig::Elt::isTableRow
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'table:table-row')) ? 1 : undef;
	}

sub	XML::Twig::Elt::isTableColumn
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'table:table-column')) ? 1 : undef;
	}

sub	XML::Twig::Elt::isTableCell
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'table:table-cell')) ? 1 : undef;
	}

sub	XML::Twig::Elt::isCovered
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name =~ /covered/)) ? 1 : undef;
	}
	
sub	XML::Twig::Elt::isSpan
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:span')) ? 1 : undef;
	}

sub	XML::Twig::Elt::isFootnoteCitation
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:footnote-citation')) ? 1 : undef;
	}

sub	XML::Twig::Elt::isFootnoteBody
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:footnote-body')) ? 1 : undef;
	}

sub	XML::Twig::Elt::isSequenceDeclarations
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:sequence-decls')) ? 1 : undef;
	}

#-----------------------------------------------------------------------------
# getText() method adaptation for complex elements
# and text output "enrichment"
# (overrides getText from OODoc::XPath)

sub	getText
	{
	my $self	= shift;
	my $element	= $self->getElement(@_);
	return undef	unless ($element && $element->isElementNode);

	my $text	= undef;
	my $begin_text	= '';
	my $end_text	= '';

	my $line_break	= $self->{'line_separator'} || '';
	if ($self->{'use_delimiters'} && $self->{'use_delimiters'} eq 'on')
		{
		my $name	= $element->getName;
		$begin_text	=
		    defined $self->{'delimiters'}{$name}{'begin'}	?
		        $self->{'delimiters'}{$name}{'begin'}		:
			($self->{'delimiters'}{'default'}{'begin'} || '');
		$end_text	=
		    defined $self->{'delimiters'}{$name}{'end'}		?
		        $self->{'delimiters'}{$name}{'end'}		:
			($self->{'delimiters'}{'default'}{'end'} || '');
		}

	$text	= $begin_text;

	if	($element->isItemList)
		{
		my $item_count = 0;
		foreach my $item ($self->getItemElementList($element))
			{
			$text .= $line_break if $item_count > 0;
			$text .= ($self->getText($item) || "");
			$item_count++;
			}
		$text .= $line_break;
		}
	elsif	(
		$element->isListItem		||
		$element->isFootnoteBody	||
		$element->isTableCell
		)
		{
		my $p = $element->getFirstChild('text:p');
		$text .= ($self->SUPER::getText($p) || '');
		}
	elsif	($element->isTable)
		{
		$text .= $self->getTableContent($element);
		}
	else
		{
		$text .= ($self->SUPER::getText($element) || '');
		}

	$text	.= $end_text;
	
	return $text;
	}

#-----------------------------------------------------------------------------
# use or don't use delimiters for flat text output

sub	outputDelimitersOn
	{
	my $self	= shift;
	$self->{'use_delimiters'}	= 'on' ;
	}

sub	outputDelimitersOff
	{
	my $self	= shift;
	$self->{'use_delimiters'}	= 'off';
	}

sub	defaultOutputTerminator
	{
	my $self	= shift;
	my $delimiter	= shift;
	$self->{'delimiters'}{'default'}{'end'} = $delimiter
		if defined $delimiter;
	return $doc->{'delimiters'}{'default'}{'end'};
	}

#-----------------------------------------------------------------------------
# setText() method adaptation for complex elements
# overrides setText from OODoc::XPath

sub	setText
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $element	= $self->getElement($path, $pos);
	return undef	unless $element;

	return $self->SUPER::setText($element, @_) if $element->isParagraph;

	my $line_break	= $self->{'line_separator'} || '';
	if	($element->isItemList)
		{
		my @text	= @_;
		foreach my $line (@text)
			{
			$self->appendItem($element, text => $line);
			}
		return wantarray ? @text : join $line_break, @text;
		}
	elsif	($element->isListItem)
		{
		return $self->setItemText($element, @_);
		}
	elsif	($element->isTableCell)
		{
		return $self->updateCell($element, @_);
		}
	else
		{
		return $self->SUPER::setText($element, @_);
		}
	}

#-----------------------------------------------------------------------------
# get the whole text content of the document in a readable (non-XML) form
# result is a list of strings or a single string

sub	getTextContent
	{
	my $self	= shift;
	return $self->selectTextContent('.*', @_);
	}

#-----------------------------------------------------------------------------
# selects headers, paragraph & list item elements matching a given pattern
# returns a list of elements
# if $action is defined, it's treated as a reference to a callback procedure
# to be executed for each node matching the pattern, with the node as arg.

sub	selectElementsByContent
	{
	my $self	= shift;
	my $pattern	= shift;
	
	my @elements	= ();
	foreach my $element ($self->getBody->getChildNodes)
		{
		next if
			(
				(! $element->isElementNode)
				||
				($element->isSequenceDeclarations)
			);
		push @elements, $element
			if (
				(! $pattern)
				||
				($pattern eq '.*')
				||
				(defined $self->_search_content
					($element, $pattern, @_, $element))
			   );
		}

	return @elements;
	}

sub	findElementsByContent	# deprecated
	{
	my $self	= shift;
	return $self->selectElementsByContent(@_);
	}

sub	replaceAll		# deprecated
	{
	my $self	= shift;
	return $self->selectElementsByContent(@_);
	}

#-----------------------------------------------------------------------------
# select the 1st element matching a given pattern

sub	selectElementByContent
	{
	my $self	= shift;
	my $pattern	= shift;
	
	foreach my $element ($self->getBody->getChildNodes)
		{
		next if
			(
				(! $element->isElementNode)
				||
				($element->isSequenceDeclarations)
			);
		return $element
			if (
				(! $pattern)
				||
				($pattern eq '.*')
				||
				(defined $self->_search_content
					($element, $pattern, @_, $element))
			   );
		}
	return undef;
	}

#-----------------------------------------------------------------------------
# selects texts matching a given pattern, with optional replacement on the fly 
# returns the whole content without pattern
# result is a list of strings or a single string

sub	selectTextContent
	{
	my $self	= shift;
	my $pattern	= shift;

	my $line_break	= $self->{'line_separator'} || '';
	my @lines	= ();
	foreach my $element ($self->getBody->getChildNodes)
		{
		next if
			(
				(! $element->isElementNode)
				||
				($element->isSequenceDeclarations)
			);
		push @lines, $self->getText($element)
			    if (
				(! $pattern)
				||
				($pattern eq '.*')
				||
				(defined $self->_search_content
					($element, $pattern, @_, $element))
			       );
		}
	return wantarray ? @lines : join $line_break, @lines;
	}

sub	findTextContent
	{
	my $self	= shift;

	$self->selectTextContent(@_);
	}

#-----------------------------------------------------------------------------
# get the list of text elements

sub	getTextElementList
	{
	my $self	= shift;
	return $self->selectChildElementsByName
			(
			$self->getBody,
			't(ext:(h|p|.*list|table.*)|able:.*)'
			);
	}

#-----------------------------------------------------------------------------
# get the list of paragraph elements

sub	getParagraphList
	{
	my $self	= shift;

	return $self->getElementList('//text:p');
	}

#-----------------------------------------------------------------------------
# get the paragraphs as a list of strings

sub	getParagraphTextList
	{
	my $self	= shift;
	
	return $self->getTextList('//text:p', @_);
	}

#-----------------------------------------------------------------------------
# get the list of header elements

sub	getHeaderList
	{
	my $self	= shift;

	return $self->getElementList('//text:h');
	}

#-----------------------------------------------------------------------------
# get the headers as a list of strings

sub	getHeaderTextList
	{
	my $self	= shift;

	return $self->getTextList('//text:h', @_);
	}

#-----------------------------------------------------------------------------
# get the list of span elements (i.e. text elements distinguished from their
# containing paragraph by any kind of attribute such as font, color, etc)

sub	getSpanList
	{
	my $self	= shift;

	return $self->getElementList('//text:span');
	}

#-----------------------------------------------------------------------------
# get the span elements as a list of strings

sub	getSpanTextList
	{
	my $self	= shift;

	return $self->getTextList('//text:span', @_);
	}

#-----------------------------------------------------------------------------
# set a span style within a text element

sub	setSpan
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= ref $path ? undef : shift;

	my $element	= undef;

	if (ref $path)
		{
		$element	= $path;
		}
	else
		{
		my $context	= shift;
		unless (ref $context)
			{
			$element = $self->getElement($path, $pos)
					or return undef;
			unshift @_, $context;
			}
		else
			{
			$element = $self->getElement
						($path, $pos, $context)
					or return undef;
			}
		}
	my $expression	= shift		or return undef;
	my $style	= shift	|| $self->{'paragraph_style'};
	my @nodes	= $element->getChildNodes;	
	NODE_LOOP : foreach my $n (@nodes)
		{
		if ($n->isElementNode)
			{
			$self->setSpan($n, $expression, $style);
			next;
			}
		next unless ($n->isTextNode);

		my $text = OpenOffice::OODoc::XPath::decode_text
						($n->getValue || "");
		while ($n && $text && ($text =~ /(.*)($expression)(.*)/))
			{
			my ($before, $selection, $after) = ($1, $2, $3);
			my $span = $self->createElement
						('text:span', $selection);
			$span->paste_before($n);
			$self->setAttribute($span, 'text:style-name', $style);
			$n->delete; $n = undef; $text = undef;
			if ($before)
				{
				$self->createTextNode($before)
						->paste_before($span);
				$text = $before;
				}
			if ($after)
				{
				$self->createTextNode($after)
						->paste_after($span);
				}
			}
		}
	}

#-----------------------------------------------------------------------------

sub	removeSpan
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= ref $path ? undef : shift;

	my $element	= ref $path ?
				$path	:
				$self->getElement($path, @_);
	return undef	unless $element;

	my $text	= "";
	my @nodes	= $element->getChildNodes;
	my $n		= undef;
	my $last_text_node = undef;
	foreach $n (@nodes)
		{
		if	($n->isTextNode)
			{
			$last_text_node	= $n;
			}
		elsif	($n->isElementNode && $n->isSpan)
			{
			my $t = $n->string_value;
			if ($last_text_node)
				{
				$last_text_node->appendText($t);
				}
			else
				{
				$last_text_node =
				    OpenOffice::OODoc::XPath::new_text_node($t);
				$element->insertBefore($last_text_node, $n);
				}
			$element->removeChild($n);
			}
		}

	return $element;
	}

#-----------------------------------------------------------------------------
# get the footnote bodies in the document

sub	getFootnoteList
	{
	my $self	= shift;
	return $self->getElementList('//text:footnote-body');
	}

#-----------------------------------------------------------------------------
# get the footnote citations in the document

sub	getFootnoteCitationList
	{
	my $self	= shift;
	return $self->getElementList('//text:footnote-citation');
	}

#-----------------------------------------------------------------------------
# get the list of tables in the document

sub	getTableList
	{
	my $self	= shift;
	return $self->getElementList('//table:table');
	}

#-----------------------------------------------------------------------------
# get a header element selected by number

sub	getHeader
	{
	my $self	= shift;
	return $self->getElement('//text:h', @_);
	}

#-----------------------------------------------------------------------------
# get the text of a header element

sub	getHeaderContent
	{
	my $self	= shift;
	return $self->getText('//text:h', @_);
	}

sub	getHeaderText
	{
	my $self	= shift;
	return $self->getText('//text:h', @_);
	}

#-----------------------------------------------------------------------------
# get the level attribute (if defined) of an element
# the level must be defined for header elements

sub	getLevel
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;

	my $element	= $self->getElement($path, $pos, @_);
	return $element->getAttribute('text:level') || "";
	}

#-----------------------------------------------------------------------------

sub	setLevel
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $level	= shift;

	my $element	= $self->getElement($path, $pos, @_);
	return $element->setAttribute('text:level' => $level);
	}

#-----------------------------------------------------------------------------
# get a paragraph element selected by number

sub	getParagraph
	{
	my $self	= shift;

	return $self->getElement('//text:p', @_);
	}

#-----------------------------------------------------------------------------
# same as getParagraph() but only between the 1st level paragraphs

sub	getTopParagraph
	{
	my $self	= shift;

	return $self->getElement('//office:body/text:p', @_);
	}

#-----------------------------------------------------------------------------
# select paragraphs by stylename

sub	selectParagraphsByStyle
	{
	my $self	= shift;

	return $self->selectElementsByAttribute
		('//text:p', 'text:style-name', shift);
	}

#-----------------------------------------------------------------------------
# select a single paragraph by stylename

sub	selectParagraphByStyle
	{
	my $self	= shift;

	return $self->selectElementByAttribute
		('//text:p', 'text:style-name', shift);
	}

#-----------------------------------------------------------------------------
# get text content of a paragraph

sub	getParagraphContent
	{
	my $self	= shift;

	return $self->getText('//text:p', @_);
	}

sub	getParagraphText
	{
	my $self	= shift;

	return $self->getText('//text:p', @_);
	}

#-----------------------------------------------------------------------------
# get ordered list root element

sub	getOrderedList
	{
	my $self	= shift;
	my $pos		= shift;

	return	(ref $pos)	?
		$pos		:
		$self->getElement('//text:ordered-list', $pos);
	}

#-----------------------------------------------------------------------------
# get unordered list root element

sub	getUnorderedList
	{
	my $self	= shift;
	my $pos		= shift;

	return	(ref $pos)	?
		$pos		:
		$self->getElement('//text:unordered-list', $pos);
	}

#-----------------------------------------------------------------------------
# get item elements list

sub	getItemElementList
	{
	my $self	= shift;
	my $list	= shift;

	return $list->children('text:list-item');
	}

#-----------------------------------------------------------------------------
# get item element text

sub	getItemText
	{
	my $self	= shift;
	my $item	= shift;

	return	undef	unless $item;
	my $para	=
		$self->selectChildElementByName($item, 'text:p');
	return	$self->getText($para);
	}

#-----------------------------------------------------------------------------
# set item element text

sub	setItemText
	{
	my $self	= shift;
	my $item	= shift;
	return	undef	unless $item;
	my $text	= shift;
	$text	= ''	unless (defined $text);

	my $para	=
		$self->selectChildElementByName($item, 'text:p');
	return	$self->setText($para, $text);
	}

#-----------------------------------------------------------------------------
# get item element style

sub	getItemStyle
	{
	my $self	= shift;
	my $item	= shift;
	return	undef	unless $item;

	my $para	=
		$self->selectChildElementByName($item, 'text:p');
	return	$self->textStyle($para);
	}

#-----------------------------------------------------------------------------
# set item element style

sub	setItemStyle
	{
	my $self	= shift;
	my $item	= shift;
	return	undef	unless $item;
	my $style	= shift;

	my $para	=
		$self->selectChildElementByName($item, 'text:p');
	return	$self->textStyle($para, $style);
	}

#-----------------------------------------------------------------------------
# append a new item in a list

sub	appendItem
	{
	my $self	= shift;
	my $list	= shift;
	return	undef	unless $list;
	my %opt		= @_;
	
	my $text	= $opt{'text'};
	my $style	= $opt{'style'};
	$style	= $opt{'attribute'}{'text:style-name'}	unless $style;
	
	unless ($style)
		{
		my $first_item	=
			$self->selectChildElementByName
				($list, 'text:list-item');
		if ($first_item)
			{
			my $p	=
				$self->selectChildElementByName
					($first_item, 'text:p');
			$style	= $self->textStyle($p)	if ($p);
			}
		}

	$style	= $self->{'paragraph_style'}	unless $style;
	my $item	= $self->appendElement($list, 'text:list-item');
	my $para	= $self->appendElement
					(
					$item, 'text:p',
					text => $text
					);
	$opt{'attribute'}{'text:style-name'} = $style;
	$self->setAttributes($para, %{$opt{'attribute'}});

	return $item;
	}

#-----------------------------------------------------------------------------
# append a new item list

sub	appendItemList
	{
	my $self	= shift;
	my %opt		= @_;
	$opt{'attribute'}{'text:style-name'} = $opt{'style'} if $opt{'style'};
	$opt{'attribute'}{'text:style-name'} = $self->{'paragraph_style'}
		unless $opt{'attribute'}{'text:style-name'};

	my $name	= 'text:unordered-list';
	if (defined $opt{'type'} && ($opt{'type'} eq 'ordered'))
		{ $name = 'text:ordered-list' ; }

	return $self->appendElement($self->getBody, $name, %opt);
	}

#-----------------------------------------------------------------------------
# insert a new item list

sub	insertItemList
	{
	my $self	= shift;
	my $path	= shift;
	my $posnode	= (ref $path)	?
				$path	:
				$self->getElement($path, shift);
	my %opt		= @_;
	$opt{'attribute'}{'text:style-name'} = $opt{'style'} if $opt{'style'};
	$opt{'attribute'}{'text:style-name'} = $self->{'paragraph_style'}
		unless $opt{'attribute'}{'text:style-name'};

	my $name	= 'text:unordered-list';
	if (defined $opt{'type'} && ($opt{'type'} eq 'ordered'))
		{ $name = 'text:ordered-list' ; }

	return $self->insertElement($posnode, $name, %opt);
	}

#-----------------------------------------------------------------------------
# row expansion utility for _expand_table
	
sub	_expand_row
	{
	my $self	= shift;
	my $row		= shift;
	unless ($row)
		{
		warn "Unknown table row\n"; return undef;
		}
	my $width	= shift || $self->{'max_cols'};

	my @cells	= $row->children('table:table-cell');

	my $cell	= undef;
	my $last_cell	= undef;
	my $rep		= 0;
	my $cellnum	= 0;
	while (@cells && ($cellnum < $width))
		{
		$cell = shift @cells; $last_cell = $cell;
		$rep  =	$cell	?
				$cell->getAttribute($COL_REPEAT_ATTRIBUTE) :
				0;
		if ($rep)
			{
			$cell->removeAttribute($COL_REPEAT_ATTRIBUTE);
			while ($rep > 1 && ($cellnum < $width))
				{
				$last_cell = $last_cell->replicateNode;
				$rep--; $cellnum++;
				}
			}
		$cellnum++ if $cell;
		}

	if ($cellnum < $width)
		{
		my $c = $self->createElement('table:table-cell');
		unless ($last_cell)
			{
			$last_cell = $c->paste_last_child($row); $rep = 0;
			}
		else
			{
			$last_cell = $c->paste_after($last_cell); $rep--;
			}
		$cellnum++;
		my $nc = $width - $cellnum;
		$last_cell = $last_cell->replicateNode($nc);
		$rep -= $nc if $rep > 0;
		}
	$last_cell->setAttribute($COL_REPEAT_ATTRIBUTE, $rep)
			if ($rep && ($rep > 1));
	
	return $row;
	}
	
#-----------------------------------------------------------------------------
# column expansion utility for _expand_table
	
sub	_expand_columns
	{
	my $self	= shift;
	my $table	= shift;
	return undef unless ($table && ref $table);
	my $width	= shift || $self->{'max_cols'};

	my @cols	= $table->children('table:table-column');

	my $col		= undef;
	my $last_col	= undef;
	my $rep		= 0;
	my $colnum	= 0;
	while (@cols && ($colnum < $width))
		{
		$col	= shift @cols; $last_col = $col;
		$rep =	$col	?
				$col->getAttribute($COL_REPEAT_ATTRIBUTE) :
				0;
		if ($rep)
			{
			$col->removeAttribute($COL_REPEAT_ATTRIBUTE);
			while ($rep > 1 && ($colnum < $width))
				{
				$last_col = $last_col->replicateNode;
				$rep--; $colnum++;
				}
			}
		$colnum++ if $col;
		}
	
	if ($colnum < $width)
		{
		my $c = $self->createElement('table:table-column');
		unless ($last_col)
			{
			$last_col = $c->paste_last_child($table); $rep = 0;
			}
		else
			{
			$last_col = $c->paste_after($last_col); $rep--;
			}
		$colnum++;
		my $nc = $width - $colnum;
		$last_col = $last_col->replicateNode($nc);
		$rep -= $nc;
		}
	$last_col->setAttribute($COL_REPEAT_ATTRIBUTE, $rep)
			if ($rep && ($rep > 1));
	return $table;
	}

#-----------------------------------------------------------------------------
# expands repeated table elements in order to address them in spreadsheets
# in the same way as in text documents
	
sub	_expand_table
	{
	my $self	= shift;
	my $table	= shift;
	my $length	= shift	|| $self->{'max_rows'};
	my $width	= shift || $self->{'max_cols'};
	return undef unless ($table && ref $table);

	$self->_expand_columns($table, $width);

	my @rows	= $table->children('table:table-row');

	my $row		= undef;
	my $last_row	= undef;
	my $rep		= 0;
	my $rownum	= 0;
	while (@rows && ($rownum < $length))
		{
		$row	= shift @rows; $last_row = $row;
		$self->_expand_row($row, $width);
		$rep =	$row	?
				$row->getAttribute($ROW_REPEAT_ATTRIBUTE) :
				0;
		if ($rep)
			{
			$row->removeAttribute($ROW_REPEAT_ATTRIBUTE);
			while ($rep > 1 && ($rownum < $length))
				{
				$last_row = $last_row->replicateNode;
				$rep--; $rownum++;
				}
			}
		$rownum++ if $row;
		}

	if ($rownum < $length)
		{
		my $r = $self->createElement('table:table-row');
		unless ($last_row)
			{
			$last_row = $r->paste_last_child($table); $rep = 0;
			}
		else
			{
			$last_row = $r->paste_after($last_row); $rep--;
			}
		$rownum++;
		$self->_expand_row($last_row, $width);
		my $nc = $length - $rownum;
		$last_row = $last_row->replicateNode($nc);
		$rep -= $nc if $rep > 0;
		}
	$last_row->setAttribute($ROW_REPEAT_ATTRIBUTE, $rep)
			if ($rep && ($rep > 1));

	return $table;
	}

#-----------------------------------------------------------------------------
# get a table size in ($lines, $columns) form

sub	getTableSize
	{
	my $self	= shift;
	my $table	= $self->getTable(shift)	or return undef;
	my $lines	= $table->children_count('table:table-row');
	my $last_row	= $self->getTableRow($table, -1) or return undef;
	my $columns	=
		$last_row->children_count('table:table-cell')	+
		$last_row->children_count('table:covered-table-cell');
	return ($lines, $columns);
	}

#-----------------------------------------------------------------------------
# get a table column descriptor element

sub	getTableColumn
	{
	my $self	= shift;
	my $p1		= shift;
	return $p1	if (ref $p1 && $p1->isTableColumn);
	my $table	= $self->getTable($p1)	or return undef;
	my $col		= shift || 0;

	return $table->child($col, 'table:table-column');
	}

sub	getColumn
	{
	my $self	= shift;
	return $self->getTableColumn(@_);
	}

#-----------------------------------------------------------------------------
# get/set a column style

sub	columnStyle
	{
	my $self	= shift;
	my $p1		= shift;
	my $column	= undef;
	if (ref $p1 && $p1->isTableColumn)
		{
		$column	= $p1;
		}
	else
		{
		$column = $self->getTableColumn($p1, shift) or return undef;
		}
	my $newstyle	= shift;

	return	defined $newstyle ?
		$self->setAttribute($column, 'table:style-name' => $newstyle)
			:
		$self->getAttribute($column, 'table:style-name');
	}

#-----------------------------------------------------------------------------
# get/set a row style

sub	rowStyle
	{
	my $self	= shift;
	my $p1		= shift;
	my $row		= undef;
	if (ref $p1 && $p1->isTableRow)
		{
		$row	= $p1;
		}
	else
		{
		$row = $self->getTableRow($p1, shift) or return undef;
		}
	my $newstyle	= shift;

	return	defined $newstyle ?
		$self->setAttribute($row, 'table:style-name' => $newstyle)
			:
		$self->getAttribute($row, 'table:style-name');
	}

#-----------------------------------------------------------------------------
# get a row element from table id and row num,
# or the row cells if wantarray

sub	getTableRow
	{
	my $self	= shift;
	my $p1		= shift;
	return $p1	if (ref $p1 && $p1->isTableRow);
	my $table	= $self->getTable($p1)	or return undef;
	my $line	= shift || 0;

	return $table->child($line, 'table:table-row');
	}

sub	getRow
	{
	my $self	= shift;
	return $self->getTableRow(@_);
	}

#-----------------------------------------------------------------------------
# get all the rows in a table

sub	getTableRows
	{
	my $self	= shift;
	my $table	= $self->getTable(shift)	or return undef;

	return $table->children('table:table-row');
	}

#-----------------------------------------------------------------------------
# spreadsheet coordinates conversion utility

sub	_coord_conversion
	{
	my $arg	= shift or return ($arg, @_);
	my $coord = uc $arg;
	return ($arg, @_) unless $coord =~ /[A-Z]/;
	
	$coord	=~ s/\s*//g;
	$coord	=~ /(^[A-Z]*)(\d*)/;
	my $c	= $1;
	my $r	= $2;
	return ($arg, @_) unless ($c && $r);
	
	my $rownum	= $r - 1;
	my @csplit	= split '', $c;
	my $colnum	= 0;
	foreach my $p (@csplit)
		{
		$colnum *= 26;
		$colnum	+= ((ord($p) - ord('A')) + 1);
		}
	$colnum--;

	return ($rownum, $colnum, @_);
	}

#-----------------------------------------------------------------------------
# get cell element by 3D coordinates ($tablenum, $line, $column)
# or by ($tablename/$tableref, $line, $column)

sub	getTableCell
	{
	my $self		= shift;
	my $p1			= shift;
	return undef	unless defined $p1;
	my $table		= undef;
	my $row			= undef;
	my $cell		= undef;

	if	(! ref $p1 || ($p1->isTable))
		{
		$table	= $self->getTable($p1)	or return undef;
		@_ = OpenOffice::OODoc::Text::_coord_conversion(@_);
		my $r	= shift || 0;
		$row	= $table->child($r, 'table:table-row')
				or return undef;
		my $c	= shift || 0;

		$cell = $row->selectChildElement
				(
				'table:(covered-|)table-cell',
				$c
				);
		}
	elsif	($p1->isTableCell)
		{
		$cell	= $p1;
		}
	else	# assume $p1 is a table row
		{
		$cell = $p1->selectChildElement
				(
				'table:(covered-|)table-cell',
				shift
				);
		}

	return ($cell && ! $cell->isCovered) ? $cell : undef;
	}

sub	getCell
	{
	my $self	= shift;
	return $self->getTableCell(@_);
	}

#-----------------------------------------------------------------------------
# get all the cells in a row

sub	getRowCells
	{
	my $self	= shift;
	my $row		= $self->getTableRow(@_)	or return undef;

	return $row->children('table:table-cell');
	}

#-----------------------------------------------------------------------------
# get table cell value

sub	getCellValue
	{
	my $self	= shift;
	my $p1		= shift;

	my $cell	= undef;
	if 	((! (ref $p1)) || $p1->isTable)
		{
		@_ = OpenOffice::OODoc::Text::_coord_conversion(@_);
		$cell = $self->getTableCell($p1, shift, shift);
		}
	elsif	($p1->isTableCell)
		{
		$cell = $p1;
		}
	elsif	($p1->isTableRow)
		{
		$cell = $self->getTableCell($p1, shift);
		}

	return undef unless $cell;

	my $cell_type	= $cell->getAttribute('table:value-type');
	if ((! $cell_type) || ($cell_type eq 'string'))		# text value
		{
		return $self->getText
			(
			$cell->first_child('text:p')
			);
		}
	else							# numeric value
		{
		return $cell->att('table:value');
		}

	return undef;
	}

#-----------------------------------------------------------------------------
# get/set a cell value type

sub	cellValueType
	{
	my $self	= shift;
	my $p1		= shift;

	my $cell	= undef;
	if 	((! (ref $p1)) || $p1->isTable)
		{
		@_ = OpenOffice::OODoc::Text::_coord_conversion(@_);
		$cell = $self->getTableCell($p1, shift, shift);
		}
	elsif	($p1->isTableCell)
		{
		$cell = $p1;
		}
	elsif	($p1->isTableRow)
		{
		$cell = $self->getTableCell($p1, shift);
		}

	return undef unless $cell;

	my $newtype	= shift;
	unless ($newtype)
		{
		return $cell->att('table:value-type');
		}
	else
		{
		if ($newtype eq 'date')
			{
			$cell->del_att('table:value');
			}
		else
			{
			$cell->del_att('table:date-value');
			}
		return $cell->set_att('table:value-type', $newtype);
		}
	}

#-----------------------------------------------------------------------------
# get/set a cell currency

sub	cellCurrency
	{
	my $self	= shift;
	my $p1		= shift;
	my $cell	= undef;
	if 	((! (ref $p1)) || $p1->isTable)
		{
		@_ = OpenOffice::OODoc::Text::_coord_conversion(@_);
		$cell = $self->getTableCell($p1, shift, shift);
		}
	elsif	($p1->isTableCell)
		{
		$cell = $p1;
		}
	elsif	($p1->isTableRow)
		{
		$cell = $self->getTableCell($p1, shift);
		}
	return undef unless $cell;

	my $newcurrency	= shift;
	unless ($newcurrency)
		{
		return $cell->att('table:currency');
		}
	else
		{
		$cell->set_att('table:value-type', 'currency');
		return $cell->set_att('table:currency', $newcurrency);
		}
	}

#-----------------------------------------------------------------------------
# get/set accessor for the formula of a table cell

sub	cellFormula
	{
	my $self	= shift;
	my $p1		= shift;
	my $cell	= undef;
	if 	((! (ref $p1)) || $p1->isTable)
		{
		@_ = OpenOffice::OODoc::Text::_coord_conversion(@_);
		$cell = $self->getTableCell($p1, shift, shift);
		}
	elsif	($p1->isTableCell)
		{
		$cell = $p1;
		}
	elsif	($p1->isTableRow)
		{
		$cell = $self->getTableCell($p1, shift);
		}
	return undef unless $cell;
	my $formula = shift;
	if (defined $formula)
		{
		if ($formula gt ' ')
			{
			$self->setAttribute($cell, 'table:formula', $formula);
			}
		else
			{
			$self->removeAttribute($cell, 'table:formula');
			}
		}
	return $self->getAttribute($cell, 'table:formula');
	}

#-----------------------------------------------------------------------------
# set value of an existing cell

sub	updateCell
	{
	my $self	= shift;
	my $p1		= shift;
	my $cell	= undef;

	if 	((! (ref $p1)) || $p1->isTable)
		{
		@_ = OpenOffice::OODoc::Text::_coord_conversion(@_);
		$cell = $self->getTableCell($p1, shift, shift);
		}
	elsif	($p1->isTableCell)
		{
		$cell = $p1;
		}
	elsif	($p1->isTableRow)
		{
		$cell = $self->getTableCell($p1, shift);
		}
	return undef	unless $cell;

	
	my $value	= shift;
	my $text	= shift;

	$text		= $value	unless defined $text;
	my $cell_type	= $cell->getAttribute('table:value-type');
	unless ($cell_type)
		{
		$cell->setAttribute('table:value-type', 'string');
		$cell_type = 'string';
		}

	my $p = $cell->first_child('text:p');
	unless ($p)
		{
		$p = $self->createParagraph($text);
		$p->paste_last_child($cell);
		}
	else
		{
		$self->SUPER::setText($p, $text);
		}

	unless ($cell_type eq 'string')
		{
		my $attribute = ($cell_type eq 'date') ?
				'table:date-value' : 'table:value';
		$cell->setAttribute($attribute, $value);
		}
	return $cell;
	}

#-----------------------------------------------------------------------------
# get/set a cell value

sub	cellValue
	{
	my $self	= shift;
	my $p1		= shift;

	my $cell	= undef;
	if 	((! (ref $p1)) || $p1->isTable)
		{
		@_ = OpenOffice::OODoc::Text::_coord_conversion(@_);
		$cell = $self->getTableCell($p1, shift, shift);
		}
	elsif	($p1->isTableCell)
		{
		$cell = $p1;
		}
	elsif	($p1->isTableRow)
		{
		$cell = $self->getTableCell($p1, shift);
		}
	return undef unless $cell;

	my $newvalue	= shift;
	unless (defined $newvalue)
		{
		return $self->getCellValue($cell);
		}
	else
		{
		return $self->updateCell($cell, $newvalue, @_);
		}
	}

#-----------------------------------------------------------------------------
# get/set a cell style

sub	cellStyle
	{
	my $self	= shift;
	my $p1		= shift;
	my $cell	= undef;
	if 	((! (ref $p1)) || $p1->isTable)
		{
		@_ = OpenOffice::OODoc::Text::_coord_conversion(@_);
		$cell = $self->getTableCell($p1, shift, shift);
		}
	elsif	($p1->isTableCell)
		{
		$cell = $p1;
		}
	elsif	($p1->isTableRow)
		{
		$cell = $self->getTableCell($p1, shift);
		}
	return undef unless $cell;

	my $newstyle	= shift;

	return defined $newstyle ?
		$self->setAttribute($cell, 'table:style-name' => $newstyle) :
		$self->getAttribute($cell, 'table:style-name');
	}

#-----------------------------------------------------------------------------
# get/set cell spanning (from a contribution by Don_Reid[at]Agilent.com)

sub	removeCellSpan
	{
	my $self	= shift;
	my $cell	= $self->getTableCell(@_) or return undef;

	my $span = $cell->getAttribute('table:number-columns-spanned') || 0;
	return undef unless ($span && $span > 0);

	$cell->removeAttribute('table:number-columns-spanned');

	my $cell_paragraph = $cell->first_child('text:p');
	my $next_cell = $cell->next_sibling;
	while ($span > 1 && $next_cell && $next_cell->isCovered)
		{
		$span--;
		$next_cell->set_name('table:table-cell');
		$next_cell->set_atts($cell->atts);
		$next_cell->del_att('table:value');
		if ($cell_paragraph)
			{
			my $p = $cell_paragraph->copy;
			$p->set_text("");
			$p->paste_first_child($next_cell);
			}
		$next_cell = $next_cell->next_sibling;
		}
	return 1;
	}

sub	cellSpan	
	{
	my $self	= shift;
	my $p1		= shift;
	my $cell	= undef;
	my $rnum	= undef;
	my $cnum	= undef;
	if 	((! (ref $p1)) || $p1->isTable)
		{
		@_ = OpenOffice::OODoc::Text::_coord_conversion(@_);
		$cell = $self->getTableCell($p1, shift, shift);
		}
	elsif	($p1->isTableCell)
		{
		$cell = $p1;
		}
	elsif	($p1->isTableRow)
		{
		$cell = $self->getTableCell($p1, shift);
		}
	return undef unless $cell;
	
	my $span = shift;	# Number of columns spanned

				# look for possible existing span
	my $old_span = $cell->getAttribute('table:number-columns-spanned')
				|| 0;
	if (! defined $span || $span == $old_span)
		{
		return $old_span;
		}
				# remove the old span
	$self->removeCellSpan($cell);
	return undef unless ($span > 1);
				# process the new span
	my $row	= $cell->getParentNode;
	my @cells = $row->children('table:table-cell');
	my $cnt = scalar(@cells);
	# which col is the current cell?
	for ($c=0; $c<$cnt; $c++) {
		if ($cell == $cells[$c]) {	# This is it
			# Check span against size!
			if (($c + $span) > $cnt) {
				$span = ($cnt - $c);
				}

			# Attach attribute to the cell,
			$cell->setAttribute('table:number-columns-spanned', 
						$span);

			# Change covered cells
			for ($i = 1; $i < $span; $i ++) {
				my $covered = $cells[$c + $i];
				my @paras = $covered->children('text:p');
				$self->replaceElement($covered, 
					'table:covered-table-cell');
				while (@paras)
					{
					my $p = shift @paras;
					$p->paste_last_child($cell) if 
						(
						defined $p->text
							&&
						$p->text ge ' '
						);
					}
				}

			last
			}
		}
	
	return $span;
	}

#-----------------------------------------------------------------------------
# get the content of a table element in a 2D array

sub	_get_row_content
	{
	my $self	= shift;
	my $row		= shift;
	
	my @row_content	= ();
	foreach my $cell ($row->children('table:table-cell'))
		{
		push @row_content, $self->getText($cell);
		}
	return @row_content;
	}

sub	getTableContent
	{
	my $self	= shift;
	my $table	= $self->getTable(shift);

	return undef	unless $table;

	my @table_content = ();
	my $headers	= $table->getFirstChild('table:table-header-rows');
	if ($headers)
		{
		push @table_content, [ $self->_get_row_content($_) ]
			for ($headers->children('table:table-row'));
		}
	push @table_content, [ $self->_get_row_content($_) ]
		for ($table->children('table:table-row'));

	if (wantarray)
		{
		return @table_content;
		}
	else
		{
		my $delimiter	= $self->{'field_separator'} || '';
		my $line_break	= $self->{'line_separator'}  || '';
		my @list	= ();
		foreach my $row (@table_content)
			{
			push @list, join($delimiter, @{$row});
			}
		return join $line_break, @list;
		}
	}

sub	getTableText
	{
	my $self	= shift;

	return $self->getTableContent(@_);
	}

#-----------------------------------------------------------------------------
# get table element selected by number

sub	getTable
	{
	my $self	= shift;
	my $table	= shift;

	return undef	unless defined $table;
	if (ref $table)
		{
		return $table->isTable ? $table : undef ;
		}
	if ($table =~ /^\d*$/)
		{
		$t = $self->getElement('//table:table', $table);
		}
	else
		{
		$t = $self->getNodeByXPath
				("//table:table[\@table:name=\"$table\"]");
		}
	my ($length, $width) = @_;
	if	(
		$length		||
			(
			$self->{'expand_tables'}		&&
			($self->{'expand_tables'} eq 'on')
			)
		)
		{
		return $self->_expand_table($t, $length, $width);
		}
	return $t;
	}

#-----------------------------------------------------------------------------
# user-controlled spreadsheet expansion

sub	normalizeSheet
	{
	my $self	= shift;
	my $table	= shift;
	unless (ref $table)
		{
		if ($table =~ /^\d*$/)
			{
			$table = $self->getElement
						('//table:table', $table);
			}
		else
			{
			$table = $self->getNodeByXPath
				("//table:table[\@table:name=\"$table\"]");
			}
		}

	unless ((ref $table) && $table->isTable)
		{
		warn	"[" . __PACKAGE__ . "::normalizeSheet] "	.
			"Missing sheet\n";
		return undef;
		}
	return $self->_expand_table($table, @_);
	}

sub	normalizeSheets
	{
	my $self	= shift;
	my $length	= shift;
	my $width	= shift;
	my @sheets	= $self->getTableList;
	my $count	= 0;
	foreach my $sheet (@sheets)
		{
		$self->normalizeSheet($sheet, $length, $width);
		$count++;
		}
	return $count;
	}

#-----------------------------------------------------------------------------
# activate/deactivate and parametrize automatic spreadsheet expansion

sub	autoSheetNormalizationOn
	{
	my $self	= shift;
	my $length	= shift || $self->{'max_rows'};
	my $width	= shift || $self->{'max_cols'};

	$self->{'expand_tables'}	= 'on';
	$self->{'max_rows'}		= $length;
	$self->{'max_cols'}		= $width;

	return 'on';
	}

sub	autoSheetNormalizationOff
	{
	my $self	= shift;
	my $length	= shift || $self->{'max_rows'};
	my $width	= shift || $self->{'max_cols'};

	$self->{'expand_tables'}	= 'no';
	$self->{'max_rows'}		= $length;
	$self->{'max_cols'}		= $width;

	return 'no';
	}

#-----------------------------------------------------------------------------
# common code for insertTable and appendTable

sub	_build_table
	{
	my $self	= shift;
	my $table	= shift;
	my $rows	= shift || $self->{'max_rows'} || 1;
	my $cols	= shift || $self->{'max_cols'} || 1;
	my %opt		=
			(
			'cell-type'	=> 'string',
			'text-style'	=> 'Table Contents',
			@_
			);

	$rows = $self->{'max_rows'} unless $rows;
	$cols = $self->{'max_cols'} unless $cols;

	my $col_proto	= $self->createElement('table:table-column');
	$self->setAttribute
		($col_proto, 'table:style-name', $opt{'column-style'})
			if $opt{'column-style'};
	$col_proto->paste_first_child($table);
	$col_proto->replicateNode($cols - 1, 'after');

	my $row_proto	= $self->createElement('table:table-row');
	my $cell_proto	= $self->createElement('table:table-cell');
	$self->cellValueType($cell_proto, $opt{'cell-type'});
	$self->cellStyle($cell_proto, $opt{'cell-style'});

	if ($opt{'paragraphs'})
		{
		my $para_proto	= $self->createElement('text:p');
		$self->setAttribute
			($para_proto, 'text:style-name', $opt{'text-style'})
				if $opt{'text-style'};
		$para_proto->paste_last_child($cell_proto);
		}

	$cell_proto->paste_first_child($row_proto);
	$cell_proto->replicateNode($cols - 1, 'after');

	$row_proto->paste_last_child($table);
	$row_proto->replicateNode($rows - 1, 'after');

	return $table;
	}

#-----------------------------------------------------------------------------
# create a new table and append it to the end of the document body (default),
# or attach it as a new child of a given element

sub	appendTable
	{
	my $self	= shift;
	my $name	= shift;
	my $rows	= shift || $self->{'max_rows'} || 1;
	my $cols	= shift || $self->{'max_cols'} || 1;
	my %opt		=
			(
			'attachment'	=> $self->getBody,
			'table-style'	=> $name,
			@_
			);

	if ($self->getTable($name))
		{
		warn	"[" . __PACKAGE__ . "::appendTable] "	.
			"Table $name exists\n";
		return	undef;
		}

	my $table = $self->appendElement
				(
				$opt{'attachment'}, 'table:table',
				attribute =>
					{
					'table:name'		=>
						$name,
					'table:style-name'	=>
						$opt{'table-style'}
					}
				)
			or return undef;

	return $self->_build_table($table, $rows, $cols, %opt);
	}

#-----------------------------------------------------------------------------

sub	insertTable
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= ref $path ? undef : shift;
	my $posnode	= $self->getElement($path, pos) or return undef;
	my $name	= shift;
	my $rows	= shift || $self->{'max_rows'} || 1;
	my $cols	= shift || $self->{'max_cols'} || 1;
	my %opt		=
			(
			'table-style'	=> $name,
			@_
			);

	if ($self->getTable($name))
		{
		warn	"[" . __PACKAGE__ . "::insertTable] "	.
			"Table $name exists\n";
		return	undef;
		}

	my $table = $self->insertElement
				(
				$posnode, 'table:table',
				attribute =>
					{
					'table:name'		=>
						$name,
					'table:style-name'	=>
						$opt{'table-style'}
					}
				)
			or return undef;

	return $self->_build_table($table, $rows, $cols, %opt);
	}

#-----------------------------------------------------------------------------

sub	renameTable
	{
	my $self	= shift;
	my $table	= $self->getTable(shift) or return undef;
	my $newname	= shift;

	if ($self->getTable($newname))
		{
		warn	"[" . __PACKAGE__ . "::renameTable] " .
			"Table name $newname already in use\n";
		return undef;
		}
	return $self->setAttribute($table, 'table:name' => $newname);
	}

#-----------------------------------------------------------------------------

sub	tableStyle
	{
	my $self	= shift;
	my $table	= $self->getTable(shift) or return undef;
	my $newstyle	= shift;

	return defined $newstyle ?
		$self->setAttribute($table, 'table:style-name' => $newstyle) :
		$self->getAttribute($table, 'table:style-name');
	}

#-----------------------------------------------------------------------------
# replicates a row in a table

sub	replicateTableRow
	{
	my $self	= shift;
	my $p1		= shift;
	my $table	= undef;
	my $row		= undef;
	if (ref $p1 && $p1->isTableRow)
		{
		$row	= $p1;
		}
	else
		{
		$table		= $self->getTable($p1) or return undef;
		my $line	= shift;
		$row 	= ($table->children('table:table-row'))[$line]
			or return undef;
		}
	my %options	=
		(
		position	=> 'after',
		@_
		);
	return $self->replicateElement($row, $row, %options);
	}

sub	replicateRow
	{
	my $self	= shift;
	return $self->replicateTableRow(@_);
	}

#-----------------------------------------------------------------------------
# replicate a row and insert the clone before (default) or after the prototype 

sub	insertTableRow
	{
	my $self	= shift;
	my $p1		= shift;
	my $row		= undef;
	if (ref $p1)
		{
		if  	($p1->isTableRow)
			{ $row = $p1; }
		else
			{
			$row = $self->getTableRow($p1, shift);
			}
		}
	else
		{
		$row = $self->getTableRow($p1, shift);
		}
	return undef	unless $row;

	my %options	=
			(
			position	=> 'before',
			@_
			);
	return $self->replicateTableRow($row, %options);
	}

sub	insertRow
	{
	my $self	= shift;
	return $self->insertTableRow(@_);
	}

#-----------------------------------------------------------------------------
# append a new row (replicating the last existing one) to a table

sub	appendTableRow
	{
	my $self	= shift;
	my $table	= shift;
	return $self->replicateTableRow($table, -1, position => 'after', @_);
	}

sub	appendRow
	{
	my $self	= shift;
	return $self->appendTableRow(@_);
	}

#-----------------------------------------------------------------------------
# append an element to the document body

sub	appendBodyElement
	{
	my $self	= shift;

	return $self->appendElement($self->getBody, @_);
	}

#-----------------------------------------------------------------------------
# create a new paragraph

sub	createParagraph
	{
	my $self	= shift;
	my $text	= shift;
	my $style	= shift;

	my $p = XML::Twig::Elt->new('text:p');
	if ($text)
		{
		$self->SUPER::setText($p, $text);
		}
	if ($style)
		{
		$self->setAttribute
				(
				$p,
				'text:style-name',
				$self->inputTextConversion($style)
				);
		}
	return $p;
	}

#-----------------------------------------------------------------------------
# add a new or existing text at the end of the document

sub	appendText
	{
	my $self	= shift;
	my $name	= shift;
	my %opt		= @_;

	my $attachment	= $opt{'attachment'} || $self->getBody;
	$opt{'attribute'}{'text:style-name'} = $opt{'style'}
			if $opt{'style'};
	unless ((ref $name) || $opt{'attribute'}{'text:style-name'})
		{
		$opt{'attribute'}{'text:style-name'} =
					$self->{'paragraph_style'};
		}

	delete $opt{'attachment'};
	delete $opt{'style'};
	return $self->appendElement($attachment, $name, %opt);
	}

#-----------------------------------------------------------------------------
# insert a new or existing text element before or after an given element

sub	insertText
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $name	= shift;
	my %opt		= @_ ;

	$opt{'attribute'}{'text:style-name'} = $opt{'style'} if $opt{'style'};

	return (ref $path)	?
		$self->insertElement($path, $name, %opt)		:
		$self->insertElement($path, $pos, $name, %opt);
	}

#-----------------------------------------------------------------------------
# create and add a new paragraph at the end of the document

sub	appendParagraph
	{
	my $self	= shift;
	my %opt		=
			(
			style		=> $self->{'paragraph_style'},
			@_
			);

	my $paragraph = $self->createParagraph($opt{'text'}, $opt{'style'});

	my $attachment	= $opt{'attachment'} || $self->getBody;
	$paragraph->paste_last_child($attachment);

	return $paragraph;
	}

#-----------------------------------------------------------------------------
# add a new header at the end of the document

sub	appendHeader
	{
	my $self	= shift;
	my %opt		=
			(
			style	=> $self->{'header_style'},
			level	=> '1',
			@_
			);

	$opt{'attribute'}{'text:level'}	= $opt{'level'};
	
	return $self->appendText('text:h',%opt);
	}

#-----------------------------------------------------------------------------
# insert a new paragraph at a given position

sub	insertParagraph
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my %opt		=
			(
			style	=> $self->{'paragraph_style'},
			@_
			);

	return (ref $path)	?
		$self->insertText($path, 'text:p', %opt)		:
		$self->insertText($path, $pos, 'text:p', %opt);
	}

#-----------------------------------------------------------------------------
# insert a new header at a given position

sub	insertHeader
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my %opt		=
			(
			style	=> $self->{'header_style'},
			level	=> '1',
			@_
			);

	$opt{'attribute'}{'text:level'}	= $opt{'level'};

	return (ref $path) ?
		$self->insertText($path, 'text:h', %opt)		:
		$self->insertText($path, $pos, 'text:h', %opt);
	}

#-----------------------------------------------------------------------------
# remove the paragraph element at a given position

sub	removeParagraph
	{
	my $self	= shift;
	my $pos		= shift;
	return $self->removeElement($pos)	if (ref $pos);
	return $self->removeElement('//text:p', $pos);
	}

#-----------------------------------------------------------------------------
# remove the header element at a given position

sub	removeHeader
	{
	my $self	= shift;
	my $pos		= shift;
	return $self->removeElement($pos)	if (ref $pos);
	return $self->removeElement('//text:h', $pos);
	}

#-----------------------------------------------------------------------------

sub	textStyle
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $element	= $self->getElement($path, $pos) or return undef;
	my $newstyle	= shift;

	if ($element->isListItem)
		{
		return defined $newstyle ?
			$self->setItemStyle($element)	:
			$self->getItemStyle($element);
		}
	else
		{
		return defined $newstyle ?
			$self->setAttribute
				($element, 'text:style-name' => $newstyle) :
			$self->getAttribute($element, 'text:style-name');
		}
	}

#-----------------------------------------------------------------------------
# deprecated methods, maintained for compatibility reasons only

sub	getStyle
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $element	= $self->getElement($path, $pos) or return undef;
	return	$self->textStyle($element);
	}

sub	setStyle
	{
	my $self	= shift;
	return	$self->textStyle(@_);
	}

#-----------------------------------------------------------------------------
1;
