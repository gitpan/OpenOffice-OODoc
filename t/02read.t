#-----------------------------------------------------------------------------
# 02read.t	OpenOffice::OODoc Installation test		(c) GENICORP
#-----------------------------------------------------------------------------

use Test;
BEGIN	{ plan tests => 9 }

use OpenOffice::OODoc	1.307;
ok($OpenOffice::OODoc::VERSION >= 1.307);

#-----------------------------------------------------------------------------

my $testfile	=	"ootest.sxw";
my $generator	=	"OpenOffice::OODoc " . $OpenOffice::OODoc::VERSION .
			" installation test";

# Opening the $testfile file
my $archive = ooFile($testfile);
unless ($archive)
	{
	ok(0); # Unable to get the $testfile file
	exit;
	}
ok(1); # Test file open

# Opening the document content
my $doc = ooDocument(archive => $archive);
unless ($doc)
	{
	ok(0); # Unable to get a regular document content
	}
else
	{
	ok(1); # Content parsed
	}

# Opening the metadata
my $meta = ooMeta(archive => $archive);
unless ($meta)
	{
	ok(0); # Unable to get regular metadata
	exit unless $doc; # Give up if neither content nor metadata
	}
else
	{
	ok(1); # Metadata parsed
	}

my $manifest = ooManifest(archive => $archive);
unless ($manifest)
	{
	ok(0); # Unable to get the manifest
	}
else
	{
	ok(1); # Manifest parsed
	}

# Checking the mime type
my $mimetype = $manifest->getMainEntry;
ok($mimetype = "application/vnd.sun.xml.writer");

# Checking the image element
ok($doc->getImageElement("Logo"));

# Selecting a paragraph by style
ok($doc->selectParagraphByStyle("Colour"));

# Checking the installation signature in the metadata
ok($meta->generator() eq $generator);

exit 0;

