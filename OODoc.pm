#-----------------------------------------------------------------------------
#	$Id : OODoc.pm 1.104 2004-03-12 JMG$
#-----------------------------------------------------------------------------

use OpenOffice::OODoc::File		1.103;
use OpenOffice::OODoc::Meta		1.002;
use OpenOffice::OODoc::Document		1.004;

#-----------------------------------------------------------------------------

package	OpenOffice::OODoc;
use 5.006_001;
our $VERSION				= 1.104;

require Exporter;
our @ISA    = qw(Exporter);
our @EXPORT = qw(ooXPath ooFile ooText ooMeta ooImage ooDocument ooStyles);

#-----------------------------------------------------------------------------
# create a common reusable XML parser in the space of the main program

sub	BEGIN
	{
	$main::XML_PARSER = XML::XPath::XMLParser->new;
	}

#-----------------------------------------------------------------------------

sub	ooFile
	{
	return OpenOffice::OODoc::File->new(@_);
	}

sub	ooXPath
	{
	return OpenOffice::OODoc::XPath->new(@_);
	}

sub	ooText
	{
	return OpenOffice::OODoc::Text->new(@_);
	}

sub	ooMeta
	{
	return OpenOffice::OODoc::Meta->new(@_);
	}

sub	ooImage
	{
	return OpenOffice::OODoc::Image->new(@_);
	}

sub	ooDocument
	{
	return OpenOffice::OODoc::Document->new(@_);
	}

sub	ooStyles
	{
	return OpenOffice::OODoc::Styles->new(@_);
	}
	
#-----------------------------------------------------------------------------
1;

=head1	NAME

OpenOffice::OODoc - A multipurpose API for OpenOffice.org document processing

=head1	SYNOPSIS

		use OpenOffice::OODoc;

			# get global access to the document
		my $doc = ooDocument(file => 'foo.sxw');
			# retrieve a paragraph matching a given content
		my $found = $doc->selectElementByContent("Dear Customer");
			# change a style
		$doc->style($found, 'Salutation') if $found;
			# insert graphics
		$doc->createImageElement
			(
			"Corporate Logo",
			style	=> "LargeLogo",
			page	=> 1,
			size	=> "3cm, 2.5cm",
			import	=> "c:\graphics\logo.png"
			);
			# append text
		$doc->appendParagraph
			(
			style	=> 'Text body',
			text	=> 'Sincerely yours'
			);
			# save the changes
		$doc->save;


=head1	DESCRIPTION

	This toolbox allows direct read/write operations on documents, without
	using the OpenOffice.org software. It provides a high-level,
	document-oriented language, and isolates the programmer from the
	details of the OpenOffice.org XML dialect and file format.

=head1	DETAILS

	The main module of the API, OpenOffice::OODoc, provides some code
	shortcuts for the programmer. So, its main function is to load the
	operational modules, i.e :

		OpenOffice::OODoc::Document
		OpenOffice::OODoc::File
		OpenOffice::OODoc::Image
		OpenOffice::OODoc::Meta
		OpenOffice::OODoc::Styles
		OpenOffice::OODoc::Text
		OpenOffice::OODoc::XPath

	For the detailed documentation, see the man pages of these modules.
	But, before using it you should read the README of the standard
	distribution, or the OpenOffice::OODoc::Intro man page, to get
	an immediate knowledge of the functionality of each one.
	Alternatively, you can download the original reference manual
	in OpenOffice.org or PDF format at

		http://www.genicorp.fr/devel/oodoc

=head2	Exported functions

	These functions are only shortcuts for the 'new' methods of the
	modules listed above. Each one is constructed from the corresponding
	module's base name, preceded by 'oo'. So, 'ooDocument' is a synonym
	of 'OpenOffice::OODoc::Document', etc.

=head1	AUTHOR/COPYRIGHT

	Initial developer: Jean-Marie Gouarne
	Copyright 2004 by Genicorp, S.A. (www.genicorp.com)
	Licensing conditions:
		- Licence Publique Generale Genicorp v1.0
		- GNU Lesser General Public License v2.1
	Contact: oodoc@genicorp.com

=cut

