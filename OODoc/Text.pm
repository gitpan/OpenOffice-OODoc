#-----------------------------------------------------------------------------
#
#	$Id : Text.pm 1.115 2005-01-29 JMG$
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
use	OpenOffice::OODoc::XPath	1.117;
our	@ISA		= qw ( OpenOffice::OODoc::XPath );
our	$VERSION	= 1.115;

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

sub	isContent
	{
	my $self	= shift;
	return $self->contentClass() ? 1 : undef;
	}

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
# text element type detection (add-in for XML::XPath)

sub	XML::XPath::Node::Element::isOrderedList
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:ordered-list')) ?  1 : undef;
	}

sub	XML::XPath::Node::Element::isUnorderedList
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:unordered-list')) ?  1 : undef;
	}

sub	XML::XPath::Node::Element::isItemList
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

sub	XML::XPath::Node::Element::isListItem
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:list-item')) ?  1 : undef;
	}

sub	XML::XPath::Node::Element::isParagraph
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:p')) ? 1 : undef;
	}

sub	XML::XPath::Node::Element::isHeader
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:h')) ? 1 : undef;
	}

sub	XML::XPath::Node::Element::headerLevel
	{
	my $element	= shift;
	return $element->getAttribute('text:level');
	}

sub	XML::XPath::Node::Element::isTable
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'table:table')) ? 1 : undef;
	}

sub	XML::XPath::Node::Element::isTableRow
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'table:table-row')) ? 1 : undef;
	}

sub	XML::XPath::Node::Element::isTableColumn
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'table:table-column')) ? 1 : undef;
	}

sub	XML::XPath::Node::Element::isTableCell
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'table:table-cell')) ? 1 : undef;
	}

sub	XML::XPath::Node::Element::isSpan
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:span')) ? 1 : undef;
	}

sub	XML::XPath::Node::Element::isFootnoteCitation
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:footnote-citation')) ? 1 : undef;
	}

sub	XML::XPath::Node::Element::isFootnoteBody
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'text:footnote-body')) ? 1 : undef;
	}

sub	XML::XPath::Node::Element::isSequenceDeclarations
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
	my $element	= $self->SUPER::getElement(@_);
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
		my $p = $self->selectChildElementByName
					($element, 'text:p');
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

#-----------------------------------------------------------------------------
# setText() method adaptation for complex elements
# overrides setText from OODoc::XPath

sub	setText
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $element	= $self->SUPER::getElement($path, $pos);
	return undef	unless $element;

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
				(defined $element->_search_content
						($pattern, @_, $element))
			   );
		}

	return @elements;
	}

sub	findElementsByContent
	{
	my $self	= shift;

	$self->selectElementsByContent(@_);
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
				(defined $element->_search_content
						($pattern, @_, $element))
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
				(defined $element->_search_content
						($pattern, @_, $element))
			       );
		}
	return wantarray ? @lines : join $line_break, @lines;
	}

sub	findTextContent
	{
	my $self	= shift;

	$self->selectTextContent(@_);
	}

#----------------------------------------------------------------------------
# replace every substring matching a given pattern in the whole text content

sub	replaceAll
	{
	my $self	= shift;
	my @list	= ();
	
	push @list, [ $self->findElementList('//text:h', @_) ];
	push @list, [ $self->findElementList('//text:p', @_) ];

	return @list;
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
			$element = $self->SUPER::getElement($path, $pos)
					or return undef;
			unshift @_, $context;
			}
		else
			{
			$element = $self->SUPER::getElement
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
			$element->insertBefore($span, $n);
			$self->setAttribute($span, 'text:style-name', $style);
			$element->removeChild($n); $n = undef; $text = undef;
			if ($before)
				{
				$n = $self->createTextNode($before);
				$element->insertBefore($n, $span);
				$text = $before;
				}
			if ($after)
				{
				my $a = $self->createTextNode($after);
				$element->insertAfter($a, $span);
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
				$self->SUPER::getElement($path, @_);
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
					XML::XPath::Node::Text->new($t);
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
	return $self->SUPER::getElement('//text:h', @_);
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

	my $element	= $self->SUPER::getElement($path, $pos, @_);
	return $element->getAttribute('text:level') || "";
	}

#-----------------------------------------------------------------------------

sub	setLevel
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $level	= shift;

	my $element	= $self->SUPER::getElement($path, $pos, @_);
	return $element->setAttribute('text:level' => $level);
	}

#-----------------------------------------------------------------------------
# get a paragraph element selected by number

sub	getParagraph
	{
	my $self	= shift;

	return $self->SUPER::getElement('//text:p', @_);
	}

#-----------------------------------------------------------------------------
# same as getParagraph() but only between the 1st level paragraphs

sub	getTopParagraph
	{
	my $self	= shift;

	return $self->SUPER::getElement('//office:body/text:p', @_);
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
		$self->SUPER::getElement('//text:ordered-list', $pos);
	}

#-----------------------------------------------------------------------------
# get unordered list root element

sub	getUnorderedList
	{
	my $self	= shift;
	my $pos		= shift;

	return	(ref $pos)	?
		$pos		:
		$self->SUPER::getElement('//text:unordered-list', $pos);
	}

#-----------------------------------------------------------------------------
# get item elements list

sub	getItemElementList
	{
	my $self	= shift;
	my $list	= shift;

	return $self->selectChildElementsByName($list, 'text:list-item');
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
				$self->SUPER::getElement($path, shift);
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

	my @cells	= $self->selectChildElementsByName
					(
					$row,
					'^table:table-cell$'
					);

	my $cell	= undef;
	my $last_cell	= $cells[0];
	my $rep		= 0;
	my $cellnum	= 0;
	while (@cells && ($cellnum < $width))
		{
		$cell = shift @cells;
		$rep  =	$cell	?
				$cell->getAttribute($COL_REPEAT_ATTRIBUTE) :
				0;
		if ($rep)
			{
			$cell->removeAttribute($COL_REPEAT_ATTRIBUTE);
			while ($rep > 1 && ($cellnum < $width))
				{
				$last_cell = $self->replicateElement
					(
					$cell, $last_cell, position => 'after'
					);
				$rep--; $cellnum++;
				}
			}
		$cellnum++ if $cell;
		}
	while ($cellnum < $width)
		{
		$last_cell = $self->appendElement($row, 'table:table-cell');
		$cellnum++;
		$rep-- if ($rep && ($rep > 0));
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
	
	my @cols	= $self->selectChildElementsByName
					(
					$table,
					'^table:table-column$'
					);

	my $col		= undef;
	my $last_col	= $cols[0];
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
				$last_col = $self->replicateElement
					(
					$col, $last_col, position => 'after'
					);
				$rep--; $colnum++;
				}
			}
		$colnum++ if $col;
		}
	
	while ($colnum < $width)
		{
		$last_col = $self->insertElement
				(
				$last_col, 'table:table-column',
				position => 'after'
				);
		$colnum++;
		$rep-- if ($rep && ($rep > 0));
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
	
	my @rows	= $self->selectChildElementsByName
					(
					$table,
					'^table:table-row$'
					);

	my $row		= undef;
	my $last_row	= $rows[0];
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
				$last_row = $self->replicateElement
					($row, $last_row, position => 'after');
				$rep--; $rownum++;
				}
			}
		$rownum++ if $row;
		}

	while ($rownum < $length)
		{
		$last_row = $self->appendElement
				(
				$table, 'table:table-row'
				);
		$self->_expand_row($last_row, $width);
		$rownum++;
		$rep-- if ($rep && ($rep > 0));
		}
	$last_row->setAttribute($ROW_REPEAT_ATTRIBUTE, $rep)
			if ($rep && ($rep > 1));

	$self->_expand_columns($table, $width);

	return $table;
	}

#-----------------------------------------------------------------------------
# get a table size in ($lines, $columns) form

sub	getTableSize
	{
	my $self	= shift;
	my $table	= $self->getTable(shift)	or return undef;
	my @rows	= $self->selectNodesByXPath
				($table, 'table:table-row');
	my $lines	= scalar @rows;
	my $last_row	= $self->getTableRow($table, -1) or return undef;
	my @cells	= $self->selectNodesByXPath
				($last_row, 'table:table-cell');
	my $columns	= scalar @cells;
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
	my $col		= shift;

	return
	  (
	  $self->selectChildElementsByName($table, 'table:table-column')
	  )[$col];
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
# get a row element from table id and row num

sub	getTableRow
	{
	my $self	= shift;
	my $p1		= shift;
	return $p1	if (ref $p1 && $p1->isTableRow);
	my $table	= $self->getTable($p1)	or return undef;
	my $line	= shift;

	return
	  (
	  $self->selectChildElementsByName($table, 'table:table-row')
	  )[$line];
	}

sub	getRow
	{
	my $self	= shift;
	return $self->getTableRow(@_);
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
		$row	=
			(
			$self->selectChildElementsByName
					($table, 'table:table-row')
			)[$r]
				or return undef;
		my $c	= shift || 0;
		$cell	=
			(
			$self->selectChildElementsByName
					($row, 'table:table-cell')
			)[$c];
		}
	elsif	($p1->isTableCell)
		{
		$cell	= $p1;
		}
	else	# assume $p1 is a table row
		{
		$cell	=
			(
			$self->selectChildElementsByName
					($p1, 'table:table-cell')
			)[shift];
		}

	return $cell;
	}

sub	getCell
	{
	my $self	= shift;
	return $self->getTableCell(@_);
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
			$self->selectChildElementByName
				($cell, 'text:p')
			);
		}
	else							# numeric value
		{
		return $cell->getAttribute('table:value');
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
		return $cell->getAttribute('table:value-type');
		}
	else
		{
		return $cell->setAttribute('table:value-type', $newtype);
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
		return $cell->getAttribute('table:currency');
		}
	else
		{
		$cell->setAttribute('table:value-type', 'currency');
		return $cell->setAttribute('table:currency', $newcurrency);
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

	my $p = $self->selectChildElementByName($cell, 'text:p');
	$p = $self->appendElement($cell, 'text:p') unless $p;
	$self->OpenOffice::OODoc::XPath::setText($p, $text);
	$cell->setAttribute('table:value', $value)
		if ($cell_type && ($cell_type ne 'string'));
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
# get/set cell spanning (contributed by Don_Reid[at]Agilent.com)

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
	if ((!defined $span) || $span < 1) {	
		return $cell->getAttribute('table:number-columns-spanned');
		}

	my $row	= $cell->getParentNode;
	my @cells = $self->selectNodesByXPath($row, 'table:table-cell');
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
				$self->replaceElement($cells[$c + $i], 
					'table:covered-table-cell');
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
	foreach my $cell
		($self->selectChildElementsByName($row, 'table:table-cell'))
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
	my $headers	= $self->selectChildElementByName
					($table, 'table:table-header-rows');
	$headers &&
	push @table_content, [ $self->_get_row_content($_) ]
		for ($self->selectChildElementsByName($headers, 'table:table-row'));
	push @table_content, [ $self->_get_row_content($_) ]
		for ($self->selectChildElementsByName($table, 'table:table-row'));

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
		$self->_expand_table($t, $length, $width);
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
			$table = $self->SUPER::getElement
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

	for (my $i = 0 ; $i < $cols ; $i++)
		{
		$self->appendElement
				(
				$table, 'table:table-column',
				attribute	=>
					{
					'table:style-name'	=>
						$opt{'column-style'}
					}
				);
		}

	for (my $r = 0 ; $r < $rows ; $r++)
		{
		my $row = $self->appendElement($table, 'table:table-row');
		for (my $c = 0 ; $c < $cols ; $c++)
			{
			my $cell = $self->appendElement
					(
					$row, 'table:table-cell',
					attribute	=>
						{
						'table:value-type'	=>
							$opt{'cell-type'},
						'table:style-name'	=>
							$opt{'cell-style'}
						}
					);
			$self->appendElement
					(
					$cell, 'text:p',
					attribute	=>
						{
						'text:style-name'	=>
							$opt{'text-style'}
						}
					);
			}
		}
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
		$row 	=
			(
			$self->selectChildElementsByName
				($table, 'table:table-row')
			)[$line]
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
			style	=> $self->{'paragraph_style'},
			@_
			);

	return $self->appendText('text:p', %opt);
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
