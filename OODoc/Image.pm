#-----------------------------------------------------------------------------
#
#	$Id : Image.pm 1.011 2004-08-03 JMG$
#
#	Initial developer: Jean-Marie Gouarne
#	Copyright 2004 by Genicorp, S.A. (www.genicorp.com)
#	Licensing conditions:
#		- Licence Publique Generale Genicorp v1.0
#		- GNU Lesser General Public License v2.1
#	Contact: oodoc@genicorp.com
#
#-----------------------------------------------------------------------------

package	OpenOffice::OODoc::Image;
use	5.006_001;
use	OpenOffice::OODoc::XPath	1.115;
use	File::Basename;
our	@ISA		= qw ( OpenOffice::OODoc::XPath );
our	$VERSION	= 1.011;

#-----------------------------------------------------------------------------
# default attributes for image style

our	%DEFAULT_IMAGE_STYLE =
	(
	references	=>
		{
		'style:family'			=> 'graphics',
		'style:parent-style-name'	=> 'Graphics'
		},
	properties	=>
		{
		'fo:clip'			=> 'rect(0cm 0cm 0cm 0cm)',
		'style:vertical-rel'		=> 'paragraph',
		'style:horizontal-rel'		=> 'paragraph',
		'style:vertical-pos'		=> 'from-top',
		'style:horizontal-pos'		=> 'from-left',
		'draw:color-mode'		=> 'standard',
		'draw:luminance'		=> '0%',
		'draw:red'			=> '0%',
		'draw:green'			=> '0%',
		'draw:blue'			=> '0%',
		'draw:gamma'			=> '1',
		'draw:color-inversion'		=> 'false'
		}
	);

#-----------------------------------------------------------------------------

sub	XML::XPath::Node::Element::isImage
	{
	my $element	= shift;
	my $name	= $element->getName;
	return ($name && ($name eq 'draw:image')) ? 1 : undef;
	}

#-----------------------------------------------------------------------------
# constructor : calling OO XPath constructor

sub	new
	{
	my $caller	= shift;
	my $class	= ref ($caller) || $caller;
	my %options	=
		(
		member		=> 'content',	# default member
		@_
		);
	my $object = $class->SUPER::new(%options);
	return	$object	?
		bless $object, $class	:
		undef;
	}

#-----------------------------------------------------------------------------
# create & insert a new image element

sub	createImageElement
	{
	my $self	= shift;
	my $name	= shift;
	my %opt		= @_;

	my $content_class = $self->contentClass;

	my $attachment	= undef;
	my $firstnode	= undef;
	my $element	= undef;
	my $description	= undef;
	my $size	= undef;
	my $position	= undef;
	my $import	= undef;
	my $path	= undef;

	if	(
			($content_class eq 'presentation')
				or
			($content_class eq 'drawing')
		)
		{
		my $target = $opt{'page'} || '';
		my $page = ref $target ?
				$target		:
				$self->selectElementByAttribute
					('draw:page', '^' . $pagename . '$');
		delete $opt{'page'};
		$path = $page;
		}
	else
		{
		$path	= $opt{'attachment'};
		}
	delete $opt{'attachment'};
	unless ($path)
		{
		$attachment	=
			($self->getElement('//office:body', 0))
			||
			($self->getElement('//style:header', 0))
			||
			($self->getElement('//style:footer', 0));
		if ($attachment && defined $opt{'page'})
			{
			$firstnode = $self->selectChildElementByName
					($attachment, 'text:(p|h)');
			}
		}
	else
		{
		$attachment = ref $path ? $path : $self->getElement($path, 0);
		}
	unless ($attachment)
		{
		warn	"[" . __PACKAGE__ .
			"::createImageElement] No valid attachment\n";
		return undef;
		}

				# parameters translation
	$opt{'draw:name'} = $name;
	if ($opt{'description'})
		{
		$description = $opt{'description'};
		delete $opt{'description'};
		}
	if ($opt{'style'})
		{
		$opt{'draw:style-name'} = $opt{'style'};
		delete $opt{'style'};
		}
	if ($opt{'size'})
		{
		$size = $opt{'size'};
		delete $opt{'size'};
		}
	if ($opt{'position'})
		{
		$position = $opt{'position'};
		$opt{'text:anchor-type'} = 'paragraph';
		delete $opt{'position'};
		}
	if ($opt{'link'})
		{
		$opt{'xlink:href'} = $opt{'link'};
		delete $opt{'link'};
		}
	if ($opt{'import'})
		{
		$import	= $opt{'import'};
		delete $opt{'import'};
		}
	
	if ($opt{'page'})	# create appropriate parameters if anchor=page
		{		# and insert before the 1st text element
		$opt{'text:anchor-type'}	= 'page';
		$opt{'text:anchor-page-number'}	= $opt{'page'};
		$opt{'draw:z-index'}		= "1";
		$element	= $firstnode ?	# is there a text element ?
			$self->insertElement	# yes, insert before it
				(
				$firstnode, 'draw:image', position => 'before'
				)
				:		# no, append to parent element
			$self->appendElement($attachment, 'draw:image');
		delete $opt{'page'};
		}
	else
		{
		if	($path)	# append to the given attachment if any
			{
			$element = $self->appendElement
				($attachment, 'draw:image');
			}
		else		# else append to a new paragraph at the end
			{
			my $p = $self->appendElement($attachment, 'text:p');
			$element = $self->appendElement($p, 'draw:image');
			}
		}

	$self->setAttributes($element, %opt);
	$self->setImageDescription($element, $description)
		if (defined $description);
	$self->setImageSize($element, $size)
		if (defined $size);
	$self->setImagePosition($element, $position)
		if (defined $position);
	$self->importImage($element, $import)
		if (defined $import);
	return $element;
	}

sub	insertImageElement
	{
	my $self	= shift;
	return $self->createImageElement(@_);
	}

#-----------------------------------------------------------------------------
# image list

sub	getImageElementList
	{
	my $self	= shift;

	return $self->getElementList('//draw:image', @_);
	}

#-----------------------------------------------------------------------------
# select an individual image element by image

sub	selectImageElementByName
	{
	my $self	= shift;
	my $text	= shift;
	return $self->selectNodeByXPath
			("//draw:image\[\@draw:name=\"$text\"\]", @_);
	}

#-----------------------------------------------------------------------------
# select a list of image elements by name

sub	selectImageElementsByName
	{
	my $self	= shift;

	return $self->selectElementsByAttribute
			('//draw:image', 'draw:name', @_);
	}

#-----------------------------------------------------------------------------
# select a list of image elements by description

sub	selectImageElementsByDescription
	{
	my $self	= shift;
	my $filter	= shift;
	my @result	= ();
	foreach my $i ($self->getImageElementList)
		{
		my $desc = $self->getXPathValue($i, 'svg:desc');
		push @result, $i if ($desc =~ /$filter/);
		}
	return @result;
	}

#-----------------------------------------------------------------------------
# select the 1st image element matching a given description

sub	selectImageElementByDescription
	{
	my $self	= shift;
	my $filter	= shift;
	my @result	= ();
	foreach my $i ($self->getImageElementList)
		{
		my $desc = $self->getXPathValue($i, 'svg:desc');
		return $i if ($desc =~ /$filter/);
		}
	return undef;
	}

#-----------------------------------------------------------------------------
# gets image element (name or ref, with type checking)

sub	getImageElement
	{
	my $self	= shift;
	my $image	= shift;
	return undef	unless $image;
	my $element	= undef;
	if (ref $image)
		{
		$element = $image;
		}
	else
		{
		$element = ($image =~ /^\//) ?
			$self->getElement($image, @_)	:
			$self->selectImageElementByName($image, @_);
		}
	return undef unless $element;
	return $element->isImage ? $element : undef;
	}

#-----------------------------------------------------------------------------
# basic image attribute accessor

sub	imageAttribute
	{
	my $self	= shift;
	my $image	= shift;
	my $attribute	= shift;
	my $value	= shift;
	my $element	= $self->getImageElement($image);
	return undef	unless $element;
	return	(defined $value)	?
		$self->setAttribute($element, $attribute => $value)	:
		$self->getAttribute($element, $attribute);
	}

#-----------------------------------------------------------------------------
# selects image element by image URL

sub	selectImageElementByLink
	{
	my $self	= shift;
	my $link	= shift;
	return $self->selectNodeByXPath
			("//draw:image\[\@xlink:href=\"$link\"\]", @_);
	}

#-----------------------------------------------------------------------------
# select image element list by image URL

sub	selectImageElementsByLink
	{
	my $self	= shift;

	return $self->selectElementsByAttribute
			('//draw:image', 'xlink:href', @_);
	}

#-----------------------------------------------------------------------------
# get/set image URL

sub	imageLink
	{
	my $self	= shift;
	return $self->imageAttribute(shift, 'xlink:href', @_);
	}

#-----------------------------------------------------------------------------
# return the internal filepath in canonical form ('Pictures/xxxx')

sub	getInternalImagePath
	{
	my $self	= shift;
	my $image	= shift;
	my $link	= $self->imageLink($image);
	if ($link && ($link =~ /^#Pictures\//))
		{
		$link =~ s/^#//;
		return $link;
		}
	else
		{
		return undef;
		}
	}

#-----------------------------------------------------------------------------
# return image coordinates

sub	getImagePosition
	{
	my $self	= shift;
	my $image	= shift;
	my $element	= $self->getImageElement($image);
	return undef	unless $element;
	my $x		= $element->getAttribute('svg:x');
	my $y		= $element->getAttribute('svg:y');
	return wantarray ? ($x, $y) : ($x . ',' . $y);
	}

#-----------------------------------------------------------------------------
# update image coordinates

sub	setImagePosition
	{
	my $self	= shift;
	my $image	= shift;
	my $element	= $self->getImageElement($image);
	return undef	unless $element;
	my ($x, $y)	= @_;
	if ($x && ($x =~ /,/))	# X and Y are concatenated in a single string
		{
		$x =~ s/\s*//g;			# remove the spaces
		$x =~ s/,(.*)//; $y = $1;	# split on the comma
		}
	$x = '0cm' unless $x; $y = '0cm' unless $y;
	$x .= 'cm' unless $x =~ /[a-zA-Z]$/;
	$y .= 'cm' unless $y =~ /[a-zA-Z]$/;
	$self->setAttributes($element, 'svg:x' => $x, 'svg:y' => $y);
	return wantarray ? ($x, $y) : ($x . ',' . $y);
	}

#-----------------------------------------------------------------------------
# get/set image coordinates

sub	imagePosition
	{
	my $self	= shift;
	my $image	= shift;
	my $x		= shift;
	my $y		= shift;

	return	(defined $x)	?
		$self->setImagePosition($image, $x, $y, @_) :
		$self->getImagePosition($image);
	}

#-----------------------------------------------------------------------------
# get image size

sub	getImageSize
	{
	my $self	= shift;
	my $image	= shift;
	my $element	= $self->getImageElement($image);
	return undef	unless $element;
	my $w		= $element->getAttribute('svg:width');
	my $h		= $element->getAttribute('svg:height');
	return wantarray ? ($w, $h) : ($w . ',' . $h);
	}

#-----------------------------------------------------------------------------
# update image size

sub	setImageSize
	{
	my $self	= shift;
	my $image	= shift;
	my $element	= $self->getImageElement($image);
	return undef	unless $element;
	my ($w, $h)	= @_;
	if ($w && ($w =~ /,/))	# W and H are concatenated in a single string
		{
		$w =~ s/\s*//g;			# remove the spaces
		$w =~ s/,(.*)//; $h = $1;	# split on the comma
		}
	$w = '0cm' unless $w; $h = '0cm' unless $h;
	$w .= 'cm' unless $w =~ /[a-zA-Z]$/;
	$h .= 'cm' unless $h =~ /[a-zA-Z]$/;
	$self->setAttributes($element, 'svg:width' => $w, 'svg:height' => $h);
	return wantarray ? ($w, $h) : ($w . ',' . $h);
	}

#-----------------------------------------------------------------------------
# get/set image size

sub	imageSize
	{
	my $self	= shift;
	my $image	= shift;
	my $w		= shift;
	my $h		= shift;

	return	(defined $w)	?
		$self->setImageSize($image, $w, $h, @_) :
		$self->getImageSize($image);
	}

#-----------------------------------------------------------------------------
# get/set image name

sub	imageName
	{
	my $self	= shift;
	return $self->imageAttribute(shift, 'draw:name', @_);
	}

#-----------------------------------------------------------------------------
# get/set image stylename

sub	imageStyle
	{
	my $self	= shift;
	return $self->imageAttribute(shift, 'draw:style-name', @_);
	}

#-----------------------------------------------------------------------------
# get image description

sub	getImageDescription
	{
	my $self	= shift;
	my $image	= shift;
	my $element	= $self->getImageElement($image);
	return $element	?
		$self->getXPathValue($element, 'svg:desc')	:
		undef;
	}

#-----------------------------------------------------------------------------
# set/update image description

sub	setImageDescription
	{
	my $self	= shift;
	my $image	= shift;
	my $element	= $self->getImageElement($image);
	return undef	unless $element;
	my $text	= shift;
	my $desc	= $self->selectChildElementByName
					($element, 'svg:desc');
	unless ($desc)
		{
		$self->appendElement($element, 'svg:desc', text => $text)
			if (defined $text);
		}
	else
		{
		if (defined $text)	{ $self->setText($desc, $text, @_);	}
		else			{ $self->removeElement($desc, @_);	}
		}

	return $desc;
	}

#-----------------------------------------------------------------------------
# delete image description

sub	removeImageDescription
	{
	my $self	= shift;
	$self->setImageDescription(shift);
	}

#-----------------------------------------------------------------------------
# get/set accessor for image description

sub	imageDescription
	{
	my $self	= shift;
	my $image	= shift;
	my $desc	= shift;
	return	(defined $desc)	?
		$self->setImageDescription($image, $desc, @_) :
		$self->getImageDescription($image, @_);
	}

#-----------------------------------------------------------------------------
# export a selected image file from OO archive

sub	exportImage
	{
	my $self	= shift;
	my $element	= $self->getImageElement(shift);
	return undef	unless $element;
	my $path	= $self->imageLink($element)	or return undef;
	unless ($path =~ /^#Pictures\//)
		{
		warn	"[" . __PACKAGE__ . "::exportImage] "		.
			"Image content $path is an external link. "	.
			"Can't be exported\n";
		return	undef;
		}
	my $target	= shift;
	unless ($target)
		{
		my $name = $self->imageName($element);
		if ($name)
			{
			$path =~ /(\..*$)/;
			$target = $name . ($1 || '');
			}
		else
			{
			$target = $path;
			}
		}
	return $self->raw_export($path, $target, @_);
	}

#-----------------------------------------------------------------------------
# export all the internal image files, or a subset of them selected by name
# return the list of exported files

sub	exportImages
	{
	my $self	= shift;
	my %opt		= @_;
	my $filter	= $opt{'filter'} || $opt{'name'} || $opt{'selection'};
	my $basename	= $opt{'path'} || $opt{'target'};
	my $suffix	= $opt{'suffix'} || $opt{'extension'};
	my $number	= defined $opt{'start_count'} ?
					$opt{'start_count'} : 1;
	my @list	= ();
	my $count	= 0;

	my @to_export	= $filter ?
				$self->selectImageElementsByName($filter, @_)
				:
				$self->getImageElementList(@_);

	IMAGE_LOOP: foreach my $image (@to_export)
		{
		my $link	= $self->imageLink($image);
		next IMAGE_LOOP unless ($link && ($link =~ /^#Pictures\//));
		my $filename	= undef;
		my $extension	= undef;
		my $target	= undef;
		if (defined $suffix)
			{
			$extension = $suffix;
			}
		else
			{
			$link =~ /(\..*$)/;
			$extension = $1 || '';
			}
		if (defined $basename)
			{
			$target = $basename . $number . $extension;
			}
		else
			{
			my $name = $self->imageName($image) || "Image$number";
			$target = $name . $extension;
			}
		$filename = $self->exportImage($image, $target);
		push @list, $filename	if $filename;
		$count++; $number++;
		}
	return wantarray ? @list : $count;
	}

#-----------------------------------------------------------------------------
# import image file

sub	importImage
	{
	my $self	= shift;
	my $element	= $self->getImageElement(shift);
	return undef	unless $element;
	my $filename	= shift;
	unless ($filename)
		{
		warn	"[" . __PACKAGE__ . "::importImage] No filename\n";
		return undef;
		}
	my ($base, $path, $suffix) =
		File::Basename::fileparse($filename, '\..*');

	my $link	= shift;
	if ($link)
		{
		$link = '#Pictures/' . $link unless $link =~ /^#Pictures\//;
		$self->imageLink($element, $link);
		}
	else
		{
		$link	= $self->imageLink($element);
		unless ($link && $link =~ /^#Pictures\//)
			{
			$link = '#Pictures/' . $base . $suffix;
			$self->imageLink($element, $link);
			}
		}
	$self->raw_import($link, $filename);	
	return $link;
	}

#-----------------------------------------------------------------------------
1;
