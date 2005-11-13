#-----------------------------------------------------------------------------
#
#	$Id : OODoc.pm 2.015 2005-11-13 JMG$
#
#	Initial developer: Jean-Marie Gouarne
#	Copyright 2004 by Genicorp, S.A. (www.genicorp.com)
#	Licensing conditions:
#		- Licence Publique Generale Genicorp v1.0
#		- GNU Lesser General Public License v2.1
#
#-----------------------------------------------------------------------------

use OpenOffice::OODoc::File		2.110;
use OpenOffice::OODoc::Meta		2.007;
use OpenOffice::OODoc::Document		2.021;
use OpenOffice::OODoc::Manifest		2.003;

#-----------------------------------------------------------------------------

package	OpenOffice::OODoc;
use 5.008_000;
our $VERSION				= 2.015;

require Exporter;
our @ISA    = qw(Exporter);
our @EXPORT = qw
	(
	ooXPath ooFile ooText ooMeta ooManifest ooImage ooDocument ooStyles
	localEncoding ooLocalEncoding ooEncodeText ooDecodeText
	ooTemplatePath workingDirectory ooWorkingDirectory
	readConfig ooReadConfig
	);

#-----------------------------------------------------------------------------
# config loader

sub	ooReadConfig
	{
	my $filename = shift;
	unless ($filename)
		{
		$filename = $INSTALLATION_PATH . '/config.xml'
			if $INSTALLATION_PATH;
		}
	unless ($filename)
		{
		warn	"[" . __PACKAGE__ . "::ooReadConfig] "	.
			"Missing configuration file\n";
		return undef;
		}
	my $config = XML::Twig->new->safe_parsefile($filename);
	unless ($config)
		{
		warn	"[" . __PACKAGE__ . "::ooReadConfig] "	.
			"Syntax error in configuration file $filename\n";
		return undef;
		}
	my $root = ($config->findnodes('//OpenOffice-OODoc'))[0];
	unless ($root && $root->isElementNode)
		{
		return undef;
		}
	foreach my $node ($root->getChildNodes)
		{
		next unless $node->isElementNode;
		my $name = $node->getName; $name =~ s/-/::/g;
		my $varname = 'OpenOffice::OODoc::' . $name;
		$$varname = $node->string_value;
		$$varname = ooDecodeText($$varname);
		}
	OpenOffice::OODoc::Styles::ooLoadColorMap();
	return 1;
	}

sub	readConfig
	{
	return ooReadConfig(@_);
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

sub	ooManifest
	{
	return OpenOffice::OODoc::Manifest->new(@_);
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

sub	ooLocalEncoding
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
		    warn	"[" . __PACKAGE__ . "::ooLocalEncoding] " .
				"Unsupported encoding\n";
		    }
		}
	return $OpenOffice::OODoc::XPath::LOCAL_CHARSET;
	}

sub	localEncoding
	{
	return ooLocalEncoding(@_);
	}
	
#-----------------------------------------------------------------------------
# accessor for default XML templates for document creation

sub	ooTemplatePath
	{
	return OpenOffice::OODoc::File::templatePath(@_);
	}

#-----------------------------------------------------------------------------
# accessor for default working directory control

sub	ooWorkingDirectory
	{
	my $path = shift;

	$OpenOffice::OODoc::File::WORKING_DIRECTORY = $path
		if defined $path;
	OpenOffice::OODoc::File::checkWorkingDirectory
		(
		$OpenOffice::OODoc::File::WORKING_DIRECTORY
		);

	return $OpenOffice::OODoc::File::WORKING_DIRECTORY;
	}

sub	workingDirectory
	{
	return ooWorkingDirectory(@_);
	}
	
#-----------------------------------------------------------------------------
# shortcuts for low-level local/utf8 code conversion 

sub	ooEncodeText
	{
	return OpenOffice::OODoc::XPath::encode_text(@_);
	}

sub	ooDecodeText
	{
	return OpenOffice::OODoc::XPath::decode_text(@_);
	}

#-----------------------------------------------------------------------------
# initialization

sub	BEGIN
	{
	my $module_path = $INC{"OpenOffice/OODoc.pm"};
	$module_path =~ s/\.pm$//;
	$OpenOffice::OODoc::INSTALLATION_PATH = $module_path;
	ooReadConfig() if ( -e "$INSTALLATION_PATH/config.xml" );
	}
#-----------------------------------------------------------------------------
1;
