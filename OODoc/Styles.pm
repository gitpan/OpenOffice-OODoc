#-----------------------------------------------------------------------------
#
#	$Id : Styles.pm 1.003 2004-03-11 JMG$
#
#	Initial developer: Jean-Marie Gouarne
#	Copyright 2004 by Genicorp, S.A. (www.genicorp.com)
#	Licensing conditions:
#		- Licence Publique Generale Genicorp v1.0
#		- GNU Lesser General Public License v2.1
#	Contact: oodoc@genicorp.com
#
#-----------------------------------------------------------------------------

package OpenOffice::OODoc::Styles;
use	5.006_001;
use	OpenOffice::OODoc::XPath	1.111;
use	File::Basename;
our	@ISA		= qw ( OpenOffice::OODoc::XPath );
our	$VERSION	= 1.003;

#-----------------------------------------------------------------------------

our	%STYLE_PATH		=
	(
	'properties'		=> 'style:properties',
	'background-image'	=> 'style:properties/style:background-image',
	'footnote-separator'	=> 'style:properties/style:footnote-sep',
	'footnote-sep'		=> 'style:properties/style:footnote-sep',
	'header'		=> 'style:header-style/style:properties',
	'footer'		=> 'style:footer-style/style:properties'
	);

#-----------------------------------------------------------------------------

sub	XML::XPath::Node::Element::isStyle
	{
	my $element	= shift;
	my $fullname	= $element->getName;
	my ($prefix, $name)	= split ':', $fullname;
	return
	 	(
		$prefix
			&&
			(($prefix eq 'style') || ($prefix eq 'number'))
			&&
			$name
			&&
			($name ne 'properties')
		)
	 	?	1 : undef;
	}

sub	XML::XPath::Node::Element::isMasterPage
	{
	my $element	= shift;
	return	(
		$element->isElementNode
			&&
		$element->getName eq 'style:master-page'
		)
		?	1 : undef;
		
	}

#-----------------------------------------------------------------------------
# constructor

sub	new
	{
	my $caller	= shift;
	my $class	= ref($caller) || $caller;
	my %options	=
		(
		member			=> 'styles',	# XML member
		@_
		);
	my $object	= $class->SUPER::new(%options);
	return	$object	?
		bless $object, $class	:
		undef;
	}

#-----------------------------------------------------------------------------
# get a particular node in a main style element

sub	getStyleNode
	{
	my $self	= shift;
	my $element	= shift;
	my $nodename	= shift;

	my $xpath	= $STYLE_PATH{$nodename} ?
				$STYLE_PATH{$nodename}	:
				$nodename;
	return $self->getNodeByXPath($element, $xpath);
	}

#-----------------------------------------------------------------------------
# create the path for a particular node in a main style element

sub	setStyleNode
	{
	my $self	= shift;
	my $element	= shift;
	my $nodename	= shift;
	my $xpath	= $STYLE_PATH{$nodename} ?
				$STYLE_PATH{$nodename}	:
				$nodename;
	return $self->makeXPath($element, $xpath);	
	}

#-----------------------------------------------------------------------------
# get named styles root element

sub	getNamedStyleRoot
	{
	my $self	= shift;
	return	$self->getElement($self->{'named_style_path'}, 0);
	}

#-----------------------------------------------------------------------------
# get automatic styles root element

sub	getAutoStyleRoot
	{
	my $self	= shift;
	return	$self->getElement($self->{'auto_style_path'}, 0);
	}

#-----------------------------------------------------------------------------
# get master styles root element

sub	getMasterStyleRoot
	{
	my $self	= shift;
	return	$self->getElement($self->{'master_style_path'}, 0);
	}

#-----------------------------------------------------------------------------
# select a list of style elements matching a given attribute, value pair
# $path may be 'auto' or 'named' to search only in automatic or named styles
# without args, returns the full style list

sub	selectStyleElementsByAttribute
	{
	my $self	= shift;
	my $attribute	= shift;
	my $value	= shift;
	my %opt		=
			(
			namespace	=> 'style',
			type		=> 'style',
			@_
			);
	my $path	= $opt{'category'};

	return	$self->getStyleList	unless ($attribute && $value);

	unless ($path)
		{
		return	($self->selectElementsByAttribute
				(
				$self->{'named_style_path'}	.
					"/$opt{'namespace'}:$opt{'type'}",
				$attribute, $value
				)
				,
			$self->selectElementsByAttribute
				(
				$self->{'auto_style_path'} 	.
					"/$opt{'namespace'}:$opt{'type'}",
				$attribute, $value
				)
			);
		}
	else
		{
		$path	= lc $path;
		if	($path =~ /^named/)
			{
			$path	= $self->{'named_style_path'};
			}
		elsif	($path =~ /^auto/)
			{
			$path	= $self->{'auto_style_path'};
			}
		else
			{
			return undef;
			}
		return	$self->selectElementsByAttribute
				(
				"$path/$opt{'namespace'}:$opt{'type'}",
				$attribute, $value
				);
		}
	}

#-----------------------------------------------------------------------------
# select a style element by name
# $path may be 'auto' or 'named' to search only in automatic or named styles

sub	selectStyleElementByAttribute
	{
	my $self	= shift;
	my $attribute	= shift;
	my $value	= shift;
	my %opt		=
			(
			namespace	=> 'style',
			type		=> 'style',
			@_
			);

	unless ($attribute)
		{
		warn	"[" . __PACKAGE__ .
			"::selectStyleElementByAttribute] Missing attribute\n";
		return undef;
		}

	my $path	= $opt{'category'};
	unless ($path)
		{
		return	$self->selectElementByAttribute
				(
				$self->{'named_style_path'} . '/style:style',
				$attribute, $value
				)
				||
			$self->selectElementByAttribute
				(
				$self->{'auto_style_path'} . '/style:style',
				$attribute, $value
				);
		}
	else
		{
		$path	= lc $path;
		if	($path =~ /^named/)
			{
			$path	= $self->{'named_style_path'};
			}
		elsif	($path =~ /^auto/)
			{
			$path	= $self->{'auto_style_path'};
			}
		else
			{
			return undef;
			}
		return	$self->selectElementByAttribute
				(
				"$path/$opt{'namespace'}:$opt{'type'}",
				$attribute, $value
				)
		}
	}

#-----------------------------------------------------------------------------

sub	selectStyleElementByName
	{
	my $self	= shift;
	return $self->selectStyleElementByAttribute('style:name', @_);
	}

#-----------------------------------------------------------------------------

sub	selectStyleElementByFamily
	{
	my $self	= shift;
	return $self->selectStyleElementByAttribute('style:family', @_);
	}

#-----------------------------------------------------------------------------

sub	selectStyleElementsByName
	{
	my $self	= shift;
	return $self->selectStyleElementsByAttribute('style:name', @_);
	}

#-----------------------------------------------------------------------------

sub	selectStyleElementsByFamily
	{
	my $self	= shift;
	return $self->selectStyleElementsByAttribute('style:family', @_);
	}

#-----------------------------------------------------------------------------
# get style element by exact name
# search for any type of style element
# parameters:
# 	path	=> <root element, or search path, default root>
# 	type	=> <style type, default 'style'>

sub	getStyleElement
	{
	my $self	= shift;
	my $style	= shift;
	return	undef	unless $style;
	return	$style->isStyle ? $style : undef	if ref $style;
	my %opt		= @_;

	my $root	= undef;
	my $type	= $opt{'type'}		|| 'style';
	my $namespace	= $opt{'namespace'}	|| 'style';

	if ($opt{'category'})
		{
		my $path	= '//office:' ;
		if	($opt{'category'} =~ /^auto/)
				{ $path .= 'automatic-styles';	}
		elsif	($opt{'category'} =~ /^named/)
				{ $path .= 'styles';		}
		else
				{ $path = $opt{'category'};		}
		$root	= $self->getElement($path, 0);
		unless ($root)
			{
			warn	"[" . __PACKAGE__ . "::getStyleElement] " .
				"Unknown search space\n";
			return undef;
			}
		}
	my $xpath	=	"//$namespace" . ':' .
				"$type\[\@style:name=\"$style\"\]";
	return $self->getNodeByXPath($xpath, $root);
	}

#-----------------------------------------------------------------------------

sub	styleName
	{
	my $self	= shift;
	my $p1		= shift;
	my $style	= undef;
	my $newname	= undef;
	if (ref $p1)
		{
		$style = $self->getStyleElement($p1) or return undef;
		$newname = shift;
		}
	else
		{
		my %opt = @_;
		$style->getStyleElement($p1, %opt) or return undef;
		$newname = $opt{'newname'};
		}
	$self->setAttribute($style, 'style:name', $newname) if $newname;
	return $self->getAttribute($style, 'style:name');
	}

#-----------------------------------------------------------------------------

sub	getAutoStyleList
	{
	my $self	= shift;
	my %opt		=
		(
		namespace	=> 'style',
		type		=> 'style',
		@_
		);
	my $path =	$self->{'auto_style_path'} . '/' .
			$opt{'namespace'} . ':' . $opt{'type'};
	return $self->getElementList($path);
	}

#-----------------------------------------------------------------------------

sub	getNamedStyleList
	{
	my $self	= shift;
	my %opt		=
		(
		namespace	=> 'style',
		type		=> 'style',
		@_
		);
	my $path =	$self->{'named_style_path'} . '/' .
			$opt{'namespace'} . ':' . $opt{'type'};
	return $self->getElementList($path);
	}

#-----------------------------------------------------------------------------

sub	getMasterStyleList
	{
	my $self	= shift;
	my %opt		=
		(
		namespace	=> 'style',
		type		=> 'master-page',
		@_
		);
	my $path =	$self->{'master_style_path'} . '/' .
			$opt{'namespace'} . ':' . $opt{'type'};
	return $self->getElementList($path);
	}

#-----------------------------------------------------------------------------

sub	getStyleList
	{
	my $self	= shift;
	return ($self->getNamedStyleList(@_), $self->getAutoStyleList(@_));
	}

#-----------------------------------------------------------------------------

sub	styleProperties
	{
	my $self	= shift;
	my $style	= shift;
	my %new_p	= @_;
	my $namespace	= $new_p{'namespace'};
	my $type	= $new_p{'type'};
	my $path	= $new_p{'path'} || $new_p{'category'};
	my $element	= $self->getStyleElement
					(
					$style,
					namespace	=> $namespace,
					type		=> $type,
					category	=> $path
					);
	return undef	unless $element;
	delete	$new_p{'namespace'};
	delete	$new_p{'type'};
	my $change	= undef;
	my $e_prefix	= $element->getPrefix;
	my $prop_name	= $e_prefix eq 'number' ?
				'number:number' : 'style:properties';
	my $properties	= $self->getChildElementByName($element, $prop_name);
	my %attr	= ();
	foreach my $k (keys %new_p)
		{
		my $a = $k =~ /:/ ? $k : $e_prefix . ':' . $k;
		$attr{$a} = $new_p{$k}; $change = 1;
		}
	if ($change)
		{
		$properties = $self->appendElement($element, $prop_name)
				unless $properties;
		$self->setAttributes($properties, %attr); 
		}
	return	$properties ? $self->getAttributes($properties) : undef;
	}

#-----------------------------------------------------------------------------

sub	getStyleAttributes
	{
	my $self	= shift;
	my $name	= shift;
	my %style	= ();
	my $element	= $self->getStyleElement($name, @_);
	unless ($element)
		{
		warn	"[" . __PACKAGE__ .
			"::getStyleAttributes] Unknown style\n";
		return %style;
		}
	%{$style{'properties'}}	= $self->styleProperties($element);
	%{$style{'references'}} = $self->getAttributes($element);
	return %style;
	}

#-----------------------------------------------------------------------------

sub	getDefaultStyleElement
	{
	my $self	= shift;
	my $style	= shift;
	if (ref $style)
		{
		return ($family->getName eq 'style:default-style') ?
			$family	: undef;
		}
	else
		{
		return $self->getNodeByXPath
		    ("style:default-style\[\@style:family=\"$style\"\]", @_);
		}
	}

#-----------------------------------------------------------------------------

sub	getDefaultStyleAttributes
	{
	my $self	= shift;
	my $style	= $self->getDefaultStyleElement(@_);
	unless ($style)
		{
		warn	"[" . __PACKAGE__ . "::getDefaultStyleAttributes] "	.
			"No available default style in the context\n";
		return undef;
		}
	return $self->getStyleAttributes($style, @_);
	}

#-----------------------------------------------------------------------------
# create a new style with given $name and %options
# by default, the style is regarded as an 'named style' if $self is
# 'styles.xml'but if $opt{path} or $opt{category} is 'auto', then
# the style is inserted as an automatic style
# if $self is a 'content.xml' object, the style is automatic

sub	createStyle
	{
	my $self	= shift;
	my $name	= shift;

	unless ($name)
		{
		warn	"[" . __PACKAGE__ . "::createStyle] "	.
			"Missing style name\n";
		return	undef;
		}
	my %opt		= @_;

	if ($self->getStyleElement($name, %opt))
		{
		warn	"[" . __PACKAGE__ . "::createStyle] "	.
			"Style $name exists\n";
		return	undef;
		}
	my $path	= undef;
	my $type	= $opt{'type'} || 'style';
	delete $opt{'type'};
	my $namespace	= $opt{'namespace'} || 'style';
	delete $opt{'namespace'};

	if	($self->getElement('//office:document-content', 0))
		{
		$path = $self->{'auto_style_path'};
		}
	elsif	($self->getElement('//office:document-styles', 0))
		{
		$path	=
			($opt{'path'}		&& $opt{'path'} =~ /auto/)
				||
			($opt{'category'}	&& $opt{'category'} =~ /auto/)
				?
			$self->{'auto_style_path'}	:
			$self->{'named_style_path'};
		}
	else
		{
		warn	"[" . __PACKAGE__ . "::createStyle] "	.
			"Style creation is not allowed in the area\n";
		return undef;
		}
	delete $opt{'path'};
	delete $opt{'category'};

	my $element	= $self->createElement($namespace . ':' . $type);
	my $attachment	= $self->getElement($path, 0);
	$attachment->appendChild($element);
	if 	($type eq 'default-style')
		{ $opt{'family'}			= $name; }
	elsif	($type eq 'number-style')
		{
		$opt{'references'}{'style:name'}	= $name;
		$opt{'family'}			= 'data-style';
		}
	else
		{ $opt{'references'}{'style:name'}	= $name; }
	$self->updateStyle($element, %opt);
	return $element;
	}

#-----------------------------------------------------------------------------
# set style attributes

sub	updateStyle
	{
	my $self	= shift;
	my $style	= shift;
	my %opt		= @_;
	my $namespace	= $opt{'namespace'};
	my $type	= $opt{'type'};
	my $path	= $opt{'path'} || $opt{'category'};
	my $element	= $self->getStyleElement
					(
					$style,
					namespace	=> $namespace,
					type		=> $type,
					category	=> $path
					);

	unless ($element)
		{
		warn	"[" . __PACKAGE__ . "::updateStyle] "	.
			"Unknown style\n";
		return undef;
		}

	if ($opt{'prototype'})
		{
		my $sv_name = $self->getAttribute($element, 'style:name');
		my %proto = $self->getStyleAttributes($opt{'prototype'});
		while (my ($key, $value) = each %proto)
			{
			if (ref $value)
				{
				while (my ($k, $v) = each %{$value})
					{
					$opt{$key}{$k} = $v
						unless $opt{$key}{$k};
					}
				}
			else
				{
				$opt{$key} = $value unless $opt{$key};
				}
			}
		delete $opt{'prototype'};
		$opt{'references'}{'style:name'} = $sv_name if $sv_name;
		}
	$opt{'references'}{'style:family'}	= $opt{'family'}
				if $opt{'family'};
	$opt{'references'}{'style:class'}	= $opt{'class'}
				if $opt{'class'};
	if ($opt{'next'})
		{
		$opt{'references'}{'style:next-style-name'} =
			ref $opt{'next'} ?
				$self->styleName($opt{'next'})	:
				$opt{'next'};
		}
	if ($opt{'parent'})
		{
		$opt{'references'}{'style:parent-style-name'} =
			ref $opt{'parent'} ?
				$self->styleName($opt{'parent'}) :
				$opt{'parent'};
		}
	$self->setAttributes($element, %{$opt{'references'}});
	$self->styleProperties($element, %{$opt{'properties'}})
				if ($opt{'properties'});

	return $self->getStyleAttributes($element);
	}

#-----------------------------------------------------------------------------
# get a page layout descriptor (pagemaster) element.
# the argument $page could be already a pagemaster, or a pageMasterStyle
# if $page appears to be a master page (or master page name), the method
# tries to get the linked page master

sub	getPageMasterElement
	{
	my $self	= shift;
	my $page	= shift;
	my $name	= undef;
	my $pagemaster	= undef;
	if (ref $page)
		{	# it is an element
		$name	= $page->getName || "";
			# is it pagemaster element ?
		if	($name eq 'style:page-master')
			{	# OK, return it
			return $page;
			}
			# is it a master page element ?
		elsif	($name eq 'style:master-page')
			{	# yes, get the page master name
			$page = $self->getAttribute
					($page, 'style:page-master-name')
				or return undef;
			}
		}
		# here we have a name
	$pagemaster = $self->selectElementByAttribute
			('//style:page-master', 'style:name', $page);
	return $pagemaster if $pagemaster;
		# it's not a page master name,
		# so we try it as a master page name
	my $masterpage = $self->selectElementByAttribute
			('//style:master-page', 'style:name', $page)
			or return undef;
		# great! we got the master page, so get the page master name
	$name	= $self->getAttribute($masterpage, 'style:page-master-name');
		# and cross the fingers
	return $self->selectElementByAttribute
			('//style:page-master', 'style:name', $name);
	}

#-----------------------------------------------------------------------------

sub	getPageMasterAttributes
	{
	my $self	= shift;
	my %attributes	= ();
	my $pagemaster	= $self->getPageMasterElement(shift);
	unless ($pagemaster)
		{
		warn	"[" . __PACKAGE__ . "::getPageMasterAttributes] " .
			"Unknown page master\n";
		return	%attributes;
		}
	
	my $node	= undef;
	%{$attributes{'references'}}	= $self->getAttributes($pagemaster);
	%{$attributes{'properties'}}	= $self->styleProperties($pagemaster);
	$node	= $self->getStyleNode($pagemaster, 'background-image');
	%{$attributes{'background-image'}} = $node ?
		$self->getAttributes($node) : ();
	$node	= $self->getStyleNode($pagemaster, 'footnote-sep');
	%{$attributes{'footnote-sep'}} = $node ?
		$self->getAttributes($node) : ();
	$node	= $self->getStyleNode($pagemaster, 'header');
	%{$attributes{'header'}} = $node ?
		$self->getAttributes($node) : ();
	$node	= $self->getStyleNode($pagemaster, 'footer');
	%{$attributes{'footer'}} = $node ?
		$self->getAttributes($node) : ();
	
	return %attributes;
	}

#-----------------------------------------------------------------------------

sub	createPageMaster
	{
	my $self	= shift;
	my $name	= shift;
	my %opt		=
			(
			category	=> 'auto',
			namespace	=> 'style',
			type		=> 'page-master',
			@_
			);
	my $pagemaster	= undef;

	if ($opt{'prototype'})
		{
		my $proto = $self->getStyleElement
				($opt{'prototype'}, type => 'page-master');
		unless ($proto)
			{
			warn	"[" . __PACKAGE__ . "::createPageMaster] " .
				"Improper prototype style\n";
			return	undef;
			}
		my $attachment	= $self->getAutoStyleRoot;
		$pagemaster = $self->replicateElement($proto, $attachment);
		$self->setAttribute($pagemaster, 'style:name', $name);
		delete $opt{'prototype'};
		}
	else
		{
		$pagemaster = $self->createStyle($name, %opt) or return undef;
		}
	
	delete $opt{'namespace'};
	delete $opt{'type'};
	delete $opt{'category'};

	$self->updatePageMaster($pagemaster, %opt);
	return $pagemaster;
	}

#-----------------------------------------------------------------------------

sub	updatePageMaster
	{
	my $self	= shift;
	my $pagemaster	= $self->getPageMasterElement(shift) or return undef;
	my %opt		= @_;
	if ($opt{'prototype'})
		{
		my $sv_name = $self->getAttribute($pagemaster, 'style:name');
		my %proto = $self->getPageMasterAttributes($opt{'prototype'});
		while (my ($key, $value) = each %proto)
			{
			if (ref $value)
				{
				while (my ($k, $v) = each %{$value})
					{
					$opt{$key}{$k} = $v
						unless $opt{$key}{$k};
					}
				}
			else
				{
				$opt{$key} = $value unless $opt{$key};
				}
			}
		delete $opt{'prototype'};
		$opt{'references'}{'style:name'} = $sv_name if $sv_name;
		}
	$self->setAttributes($pagemaster, %{$opt{'references'}});
	delete $opt{'references'};
	$self->styleProperties($pagemaster, %{$opt{'properties'}});
	delete $opt{'properties'};
	my %p		= ();
	$p{'background-image'}	=
		$self->setStyleNode($pagemaster, 'background-image');
	$p{'footnote-sep'}	=
		$self->setStyleNode($pagemaster, 'footnote-sep');
	$p{'header'}		=
		$self->setStyleNode($pagemaster, 'header');
	$p{'footer'}		=
		$self->setStyleNode($pagemaster, 'footer');

	foreach my $k (keys %opt)
		{
		my $node = $p{$k} or next;
		my %parm = %{$opt{$k}}; my %attr = ();
		foreach my $name (keys %parm)
			{
			if	($name eq 'link')
				{
				$attr{'xlink:href'} = $parm{'link'};
				}
			elsif	(! ($name =~ /:/))
				{
				$attr{"style:$name"} = $parm{$name};
				}
			else
				{
				$attr{$name} = $parm{$name};
				}
			}
		$self->setAttributes($node, %attr);
		}

	return $self->getPageMasterAttributes($pagemaster);
	}

#-----------------------------------------------------------------------------
# switch page orientation (portrait -> landscape or landscape -> portrait)

sub	switchPageOrientation
	{
	my $self	= shift;
	my $page	= $self->getPageMasterElement(shift);
	my %op		= $self->styleProperties($page);
	my %np		= ();
	$np{'fo:page-width'}	= $op{'fo:page-height'};
	$np{'fo:page-height'}	= $op{'fo:page-width'};
	my $o		= $op{'style:print-orientation'};
	if ($o)
		{
		if	($o eq 'portrait')
			{
			$np{'style:print-orientation'} = 'landscape';
			}
		elsif	($o eq 'landscape')
			{
			$np{'style:print-orientation'} = 'portrait';
			}
		}
	return $self->styleProperties($page, %np);
	}

#-----------------------------------------------------------------------------
# get the page content for a given page style

sub	getMasterPageElement
	{
	my $self	= shift;
	my $name	= shift;
	if (ref $name)
		{
		return	$name->getName eq 'style:master-page'	?
			$name : undef;
		}
	else
		{
		return $self->selectElementByAttribute
				('//style:master-page', 'style:name', $name);
		}
	}

#-----------------------------------------------------------------------------
# get/set the page master name of a given master page

sub	pageMasterStyle
	{
	my $self	= shift;
	my $masterpage	= $self->getMasterPageElement(shift) or return undef;
	my $pagemaster	= shift;
	unless ($pagemaster)
		{
		return $self->getAttribute
				($masterpage, 'style:page-master-name');
		}
	else
		{
		my $pm_name = ref $pagemaster ?
			$pm_name = $self->getAttribute
					($pagemaster, 'style:name')	:
			$pagemaster;
		$self->setAttribute
			($masterpage, 'style:page-master-name' => $pm_name);
		return $pm_name;
		}
	}

#-----------------------------------------------------------------------------
# get the background image node in a given page master

sub	getBackgroundImageElement
	{
	my $self	= shift;
	my $page	= shift;
	my $pagemaster	= $self->getPageMasterElement($page);
	unless ($pagemaster)
		{
		my $masterpage = $self->getMasterPageElement($page)
			or return undef;
		my $name = $self->pageMasterStyle($masterpage);
		$pagemaster	= $self->getPageMasterElement($name)
			or return undef;
		}
	return	$self->getStyleNode($pagemaster, 'background-image');
	}

#-----------------------------------------------------------------------------
# get/set a background image link

sub	backgroundImageLink
	{
	my $self	= shift;
	my $page	= shift;
	my $pagemaster	= $self->getPageMasterElement($page);
	unless ($pagemaster)
		{
		my $masterpage = $self->getMasterPageElement($page)
			or return undef;
		my $name = $self->pageMasterStyle($masterpage);
		$pagemaster	= $self->getPageMasterElement($name)
			or return undef;
		}
	my $newlink = shift;
	my $node = $self->getStyleNode($pagemaster, 'background-image');
	unless (defined $newlink)
		{
		return $node ?
			$self->getAttribute($node, 'xlink:href')	:
			undef;
		}
	else
		{
		my $xpath =	$STYLE_PATH{'background-image'}		.
				'[@xlink:href="' . $newlink . '"]';
		return $self->makeXPath($pagemaster, $xpath);
		}
	}

#-----------------------------------------------------------------------------

sub	getBackgroundImageAttributes
	{
	my $self	= shift;
	my $node	= $self->getBackgroundImageElement(@_)
				or return undef;
	return $self->getAttributes($node);
	}

#-----------------------------------------------------------------------------
# create or update a backgound image element associated to a given pagemaster

sub	setBackgroundImage
	{
	my $self	= shift;
	my $page	= shift;
	my $pagemaster	= $self->getPageMasterElement($page);
	unless ($pagemaster)
		{
		my $masterpage = $self->getMasterPageElement($page)
			or return undef;
		my $name = $self->pageMasterStyle($masterpage);
		$pagemaster	= $self->getPageMasterElement($name)
			or return undef;
		}
	my %opt		=
			(
			'style:position'	=> 'center center',
			'style:repeat'		=> 'no-repeat',
			'xlink:type'		=> 'simple',
			'xlink:actuate'		=> 'onLoad',
			@_
			);
	
	my $node = $self->makeXPath
				(
				$pagemaster,
				$STYLE_PATH{'background-image'}
				)
				or return undef;
	if ($opt{'link'})
		{
		$opt{'xlink:href'}	= $opt{'link'};
		delete $opt{'link'};
		}
	if ($opt{'import'})
		{
		$self->importBackgroundImage($pagemaster, $opt{'import'});
		delete $opt{'import'};
		}
	$self->setAttributes($node, %opt);
	return $node;
	}

#-----------------------------------------------------------------------------

sub	exportBackgroundImage
	{
	my $self	= shift;
	my $source	= $self->backgroundImageLink(shift)
				or return undef;
	$self->raw_export($source, @_);
	}

#-----------------------------------------------------------------------------

sub	importBackgroundImage
	{
	my $self	= shift;
	my $page	= shift;
	my $pagemaster	= $self->getPageMasterElement($page);
	unless ($pagemaster)
		{
		my $masterpage = $self->getMasterPageElement($page)
			or return undef;
		my $name = $self->pageMasterStyle($masterpage);
		$pagemaster	= $self->getPageMasterElement($name)
			or return undef;
		}
	my $filename	= shift;
	unless ($filename)
		{
		warn	"[" . __PACKAGE__ . "::importBackgroundImage] "	.
			"No source file name\n";
		return undef;
		}
	my ($base, $path, $suffix) =
		File::Basename::fileparse($filename, '\..*');

	my $link	= shift;
	if ($link)
		{
		$link = '#Pictures/' . $link unless $link =~ /^#Pictures\//;
		$self->backgroundImageLink($pagemaster, $link);
		}
	else
		{
		$link	= $self->backgroundImageLink($pagemaster);
		unless ($link && $link =~ /^#Pictures\//)
			{
			$link = '#Pictures/' . $base . $suffix;
			$self->backgroundImageLink($pagemaster, $link);
			}
		}
	$self->raw_import($link, $filename);	
	return $link;
	}

#-----------------------------------------------------------------------------

sub	createMasterPage
	{
	my $self	= shift;
	my $name	= shift;
	my $element	= $self->getMasterPageElement($name);
	if ($element)
		{
		warn	"[" . __PACKAGE__ . "::createMasterPage] "	.
			"Master page $name exists\n";
		return	undef;
		}
	my %opt		= @_;
	my $root	= $self->getElement('//office:master-styles', 0);
	unless ($root)
		{
		warn	"[" . __PACKAGE__ . "::createMasterPage] "	.
			"No master styles space in the document\n";
		return	undef;
		}

	$opt{'style:name'} = $name;
	if ($opt{'page-master'})
		{
		$opt{'style:page-master-name'}	= $opt{'page-master'};
		delete $opt{'page-master'};
		}
	if ($opt{'next'})
		{
		$opt{'style:next-style-name'}	= $opt{'next'};
		delete $opt{'next'};
		}
	return $self->appendElement
				(
				$root,
				'style:master-page',
				attribute	=> { %opt }
				);
	}

#-----------------------------------------------------------------------------

sub	masterPageHeader
	{
	my $self	= shift;
	my $masterpage	= $self->getMasterPageElement(shift) or return undef;
	my $element	= shift;
	unless ($element)
		{
		return $self->getNodeByXPath($masterpage, '/style:header');
		}
	else
		{
		my $node = $self->makeXPath($masterpage, '/style:header');
		return $self->appendElement($node, $element, @_);
		}
	}

#-----------------------------------------------------------------------------

sub	masterPageFooter
	{
	my $self	= shift;
	my $masterpage	= $self->getMasterPageElement(shift) or return undef;
	my $element	= shift;
	unless ($element)
		{
		return $self->getNodeByXPath($masterpage, '/style:footer');
		}
	else
		{
		my $node = $self->makeXPath($masterpage, '/style:footer');
		return $self->appendElement($node, $element, @_);
		}
	}

#-----------------------------------------------------------------------------

sub	getHeaderParagraph
	{
	my $self	= shift;
	my $root	= $self->masterPageHeader(shift) or return undef;
	my $n		= shift;
	return $self->getElement('text:p', $n, $root);
	}

#-----------------------------------------------------------------------------

sub	getFooterParagraph
	{
	my $self	= shift;
	my $root	= $self->masterPageFooter(shift) or return undef;
	my $n		= shift;
	return $self->getElement('text:p', $n, $root);
	}

#-----------------------------------------------------------------------------

sub	updateDefaultStyle
	{
	my $self	= shift;
	my $style	= $self->getDefaultStyleElement(shift);
	unless ($style)
		{
		warn	"[" . __PACKAGE__ . "::updateDefaultStyle] "	.
			"Unavailable default style in the context\n";
		return undef;
		}
	return $self->updateStyle($style, @_);
	}

#-----------------------------------------------------------------------------
# remove a given style element (with element type checking)

sub	removeStyle
	{
	my $self	= shift;
	my $element	= $self->getStyleElement(@_);
	if ($element && $element->isStyle)
		{
		return $self->removeElement($element);
		}
	else
		{
		warn	"[" . __PACKAGE__ . "::removeStyle] "	.
			"Unknown style or non-style element\n";
		return undef;
		}
	}

#-----------------------------------------------------------------------------
1;
