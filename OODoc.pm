#-----------------------------------------------------------------------------
#	$Id : OODoc.pm 1.106 2004-05-27 JMG$
#-----------------------------------------------------------------------------

use OpenOffice::OODoc::File		1.103;
use OpenOffice::OODoc::Meta		1.003;
use OpenOffice::OODoc::Document		1.005;

#-----------------------------------------------------------------------------

package	OpenOffice::OODoc;
use 5.008_000;
our $VERSION				= 1.106;

require Exporter;
our @ISA    = qw(Exporter);
our @EXPORT = qw
	(
	ooXPath ooFile ooText ooMeta ooImage ooDocument ooStyles
	localEncoding
	);

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
# accessor for local character set control

sub	localEncoding
	{
	my $newcharset = shift;
	if ($newcharset)
	    	{
	    	if (Encode::find_encoding($newcharset))
		    {
		    $OpenOffice::OODoc::XPath::LOCAL_CHARSET = $newcharset;
		    }
		else
		    {
		    warn	"[" . __PACKAGE__ . "::localEncoding] " .
				"Unsupported encoding\n";
		    }
		}
	return $OpenOffice::OODoc::XPath::LOCAL_CHARSET;
	}

#-----------------------------------------------------------------------------
1;

=head1	NAME

OpenOffice::OODoc - A library for direct OpenOffice.org document processing

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

The detailed documentation is organised on a by-module basis.
There is a man page for each one in the list above.
But, before using it you should read the README of the standard
distribution, or the OpenOffice::OODoc::Intro man page, to get
an immediate knowledge of the functionality of each one.
Alternatively, you can download the original reference manual
in OpenOffice.org or PDF format at http://www.genicorp.com/devel/oodoc

=head2	Exported functions

Every "ooXxx" function below is only a shortcut for the constructor
("new") in a submodule of the API. See the man page of the
corresponding module for details.

=head3	localEncoding

	Accessor to get/set the user's local character set
	(see $OpenOffice::OODoc::XPath::LOCAL_CHARSET in the
	OpenOffice::OODoc::XPath man page).

	Example:

		$old_charset = localEncoding();
		localEncoding('iso-8859-15');

	If the given argument is an unsupported encoding, an error
	message is produced and the old encoding is preserved. So
	this accessor is safer than a direct update of the
	$OpenOffice::OODoc::XPath::LOCAL_CHARSET variable.

	The default local character set is "iso-8859-1".
	Should be set to the appropriate value by the application
	before processing.

	See the Encode::Supported (Perl) documentation for the list
	of supported encodings.

=head3	ooDocument

	Shortcut for OpenOffice::OODoc::Document->new

=head3	ooFile

	Shortcut for OpenOffice::OODoc::File->new

=head3	ooImage

	Shortcut for OpenOffice::OODoc::Image->new

=head3	ooStyles

	Shortcut for OpenOffice::OODoc::Styles->new

=head3	ooText

	Shortcut for OpenOffice::OODoc::Text->new

=head3	ooXPath

	Shortcut for OpenOffice::OODoc::XPath->new

=head2	Special variable

	$XML_PARSER is a reserved variable in the space of the
	main program. It contains a reusable XML Parser
	(XML::XPath::XMLParser object), automatically created.
	Advanced, XPath-aware applications may reuse this parser
	(see the documentation of the XML::XPath Perl module) but
	they must *NOT* set the variable.

=head1	AUTHOR/COPYRIGHT

Initial developer: Jean-Marie Gouarne

Copyright 2004 by Genicorp, S.A. (http://www.genicorp.com)

Licensing conditions:

	- Licence Publique Generale Genicorp v1.0
	- GNU Lesser General Public License v2.1

Contact: oodoc@genicorp.com

=cut

