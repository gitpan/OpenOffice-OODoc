#-----------------------------------------------------------------------------
#
#	$Id : Meta.pm 1.002 2004-03-07 JMG$		(c) GENICORP 2004
#
#	Initial developer: Jean-Marie Gouarne
#	Copyright 2004 by Genicorp, S.A. (www.genicorp.com)
#	Licensing conditions:
#		- Licence Publique Generale Genicorp v1.0
#		- GNU Lesser General Public License v2.1
#	Contact: oodoc@genicorp.com
#
#-----------------------------------------------------------------------------

package	OpenOffice::OODoc::Meta;
use	5.006_001;
use	OpenOffice::OODoc::XPath	1.111;
our	@ISA		= qw ( OpenOffice::OODoc::XPath );
our	$VERSION	= 1.002;

#-----------------------------------------------------------------------------
# constructor : calling OOXPath constructor with 'meta' as member choice

sub	new
	{
	my $caller	= shift;
	my $class	= ref ($caller) || $caller;
	my %options	=
		(
		utf8		=> 1,
		member		=> 'meta',
		body_path	=> '//office:meta',
		@_
		);

	my $object = $class->SUPER::new(%options);
	if ($object)	{ return bless $object, $class; }
	else		{ return undef; }
	}

#-----------------------------------------------------------------------------
# generic read/write accessor for text elements

sub	accessor
	{
	my ($self, $path, $value)	= @_;

	my $element	= $self->getElement($path, 0);
	unless ($element)
		{
		return undef	unless defined $value;
		my $name	= $path;
		$name	=~ s/\/*//g;
		$element = $self->appendElement
				($self->getBody, $name, text => $value);
		return $value;
		}

	return 	(defined $value)			?
		$self->setText($element, $value)	:
		$self->getText($element);
	}

#-----------------------------------------------------------------------------
# get/set the 'generator' field (i.e. the signature of the office software)

sub	generator
	{
	my $self	= shift;
	return $self->accessor('//meta:generator', @_);
	}

#-----------------------------------------------------------------------------
# get/set the 'title' field

sub	title
	{
	my $self	= shift;
	return $self->accessor('//dc:title', @_);
	}

#-----------------------------------------------------------------------------
# get/set the 'description' field

sub	description
	{
	my $self	= shift;
	return $self->accessor('//dc:description', @_);
	}

#-----------------------------------------------------------------------------
# get/set the 'subject' field

sub	subject
	{
	my $self	= shift;
	return $self->accessor('//dc:subject', @_);
	}

#-----------------------------------------------------------------------------
# get/set the 'creation-date' field
# in OpenOffice.org date format

sub	creation_date
	{
	my $self	= shift;
	return $self->accessor('//meta:creation-date', @_);
	}

#-----------------------------------------------------------------------------
# get/set the 'creator' field (i.e. author)

sub	creator
	{
	my $self	= shift;
	return $self->accessor('//dc:creator', @_);
	}

#-----------------------------------------------------------------------------
# get/set the 'date' (i.e. the date of last update) field
# in OpenOffice.org (normally ISO-8601) date format

sub	date
	{
	my $self	= shift;
	return $self->accessor('//dc:date', @_);
	}

#-----------------------------------------------------------------------------
# get/set the 'language' code (ex : 'fr-FR') of the document

sub	language
	{
	my $self	= shift;
	return $self->accessor('//dc:language', @_);
	}

#-----------------------------------------------------------------------------
# get/set the 'editing-cycles' field (i.e. the number of editing sessions)

sub	editing_cycles
	{
	my $self	= shift;
	return $self->accessor('//meta:editing-cycles', @_);
	}

#-----------------------------------------------------------------------------
# get/set the 'editing-duration' field (i.e. the total elapsed time for all
# the editing sessions) in OpenOffice.org (ISO-8601) format

sub	editing_duration
	{
	my $self	= shift;
	return $self->accessor('//meta:editing-duration', @_);
	}

#-----------------------------------------------------------------------------
# get/set the 'keywords' list
# optional arguments are a list of one or more keywords ;
# each argument is appended to the keywords list only if not present in
# the old list ;
# result is an array of strings or a single string (with ',' as keyword
# separator) ; optional newly added keywords are included in the result set

sub	keywords
	{
	my $self	= shift;
	my @new_words	= @_;

	$self->addKeyword($_)	for @new_words;
	
	my $base	= $self->getElement('//meta:keywords', 0);
	return undef	unless $base;

	my @list	= ();

	foreach my $element
		($self->selectChildElementsByName($base, 'meta:keyword'))
		{ push @list, $self->getText($element); }

	return wantarray ? @list : join ", ", @list;
	}

#-----------------------------------------------------------------------------
# append a new keyword (if unknown) in the keywords list

sub	addKeyword
	{
	my $self	= shift;
	my $new_word	= shift;

	my $kw_base	= $self->getElement('//meta:keywords', 0);
	if ($kw_base)
	    {
	    foreach my $element
		($self->selectChildElementsByName($kw_base, 'meta:keyword'))
		{
		my $old_word	= $self->getText($element);
		return undef	if ($old_word eq $new_word);
		}
	    }
	else
	    {
	    $kw_base	=
		$self->appendElement('//office:meta', 0, 'meta:keywords');
	    }

	$self->appendElement($kw_base, 'meta:keyword', text => $new_word);
	return $new_word;
	}

#-----------------------------------------------------------------------------
# remove a given keyword (if known) from the keyword list

sub	removeKeyword
	{
	my $self	= shift;
	my $word	= shift;

	my $kw_base	= $self->getElement('//meta:keywords', 0);
	return undef	unless $kw_base;
	
	foreach my $element
		($self->selectChildElementsByName($kw_base, 'meta:keyword'))
		{
		my $old_word	= $self->getText($element);
		$kw_base->removeChild($element)	if ($old_word =~ /$word/);
		}
	}
	
#-----------------------------------------------------------------------------
# remove the keyword list

sub	removeKeywords
	{
	my $self	= shift;
	return $self->removeElement('//meta:keywords', 0);
	}

#-----------------------------------------------------------------------------
# get/set the list of the 4 user defined fields of the meta-data
# without argument, returns a hash where keys are the field names
# and values are the field values
# to set/update the user-defined fields, arguments should be passed
# as a hash with the same structure

sub	user_defined
	{
	my $self	= shift;
	my %new_fields	= @_;

	my @elements	= $self->getElementList('//meta:user-defined');

	if (%new_fields)
		{
		my $count	= 0;
		foreach my $key (sort keys %new_fields)
			{
			my $element	= $elements[$count];
			last	unless $element;
			$self->setAttribute($element, 'meta:name', $key);
			$self->setText($element, $new_fields{$key});
			$count++;
			last if $count > 3;
			}
		}
	
	my %fields	= ();
	foreach my $element (@elements)
		{
		my $name	= $self->getAttribute($element, 'meta:name');
		my $content	= $self->getText($element);
		$fields{$name}	= $content;
		}

	return %fields;
	}

#-----------------------------------------------------------------------------
# get/set the 'statistic' element of the meta-data
# result is a hash where keys are OpenOffice variable names and values are
# the current values of these variables in meta.xml ;
# to update all or part of the statistic fields, you should pass a hash
# with the same structure
# WARNING : you should not update statistics unless you know exactly what
# you do ; updating statistic attributes may introduce inconsistencies between
# the meta-data and the real content of the document

sub	statistic
	{
	my $self	= shift;
	my %new_fields	= @_;

	my $element	= $self->getElement('//meta:document-statistic', 0);

	unless (%new_fields)
		{ return $self->getAttributes($element);		}
	else
		{ return $self->setAttributes($element, %new_fields);	}
	
	}

#-----------------------------------------------------------------------------
1;

=head1	NAME

OpenOffice::OODoc::Meta - Interface for metadata manipulation

=cut
