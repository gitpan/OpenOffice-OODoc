#-----------------------------------------------------------------------------
#
#	$Id : OODoc.pm 1.103 2004-03-07 JMG$		(c) GENICORP 2004
#
#	Initial developer: Jean-Marie Gouarne
#	Copyright 2004 by Genicorp, S.A. (www.genicorp.com)
#	Licensing conditions:
#		- Licence Publique Generale Genicorp v1.0
#		- GNU Lesser General Public License v2.1
#	Contact: oodoc@genicorp.com
#
#	Main module for access to OpenOffice.org documents
#
#-----------------------------------------------------------------------------

use OpenOffice::OODoc::File		1.103;
use OpenOffice::OODoc::Meta		1.002;
use OpenOffice::OODoc::Document		1.003;

#-----------------------------------------------------------------------------

package	OpenOffice::OODoc;
use 5.006_001;
our $VERSION				= 1.103;

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

__END__

=head1 NAME

OpenOffice::OODoc - A multipurpose API for OpenOffice.org document processing

=head1 SYNOPSIS

This module allows direct read/write operations on documents, without using
the OpenOffice.org software. It provides a high-level, document-oriented
language, and isolates the programmer from the details of the OpenOffice.org
XML dialect and file format.

A full reference manual is available at http://www.genicorp.fr/devel/oodoc
in OpenOffice.org (SXW) format. For a short introduction, please look at the
README file in the distribution.

=cut

