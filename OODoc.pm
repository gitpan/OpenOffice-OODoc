#-----------------------------------------------------------------------------
#
#	$Id : OODoc.pm 1.102 2003-11-05 JMG$		(c) GENICORP 2003
#
#	Initial developer: Jean-Marie Gouarne
#	Copyright 2003 by Genicorp, S.A. (www.genicorp.com)
#	Licensing conditions:
#		- Licence Publique Generale Genicorp v1.0
#		- GNU Lesser General Public License v2.1
#	Contact: oodoc@genicorp.com
#
#	Main module for access to OpenOffice.org documents
#
#-----------------------------------------------------------------------------

use OpenOffice::OODoc::File		1.102;
use OpenOffice::OODoc::Meta		1.001;
use OpenOffice::OODoc::Document		1.002;

#-----------------------------------------------------------------------------

package	OpenOffice::OODoc;
use 5.006_001;
our $VERSION				= 1.102;

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
