#-----------------------------------------------------------------------------
# 02read.t	OpenOffice::OODoc 1.105 Installation test	(c) GENICORP
#-----------------------------------------------------------------------------

use Test;
BEGIN	{ plan tests => 7 }

use OpenOffice::OODoc	1.105;
ok(1);

#-----------------------------------------------------------------------------

my $testfile = "ootest.sxw";

# Opening the $testfile file
my $archive = ooFile($testfile)
	or die "Unable to get the $testfile file\n";
ok(1);

# Opening the document content
my $doc = ooDocument(archive => $archive)
	or die "Unable to get a regular OpenOffice.org document content\n";
ok(1);

# Opening the metadata
my $meta = ooMeta(archive => $archive)
	or die "Unable to get regular OpenOffice.org metadata\n";
ok(1);

# Checking the image element
ok($doc->getImageElement("Logo"));

# Selecting a paragraph by style
ok($doc->selectParagraphByStyle("Colour"));

# Checking the installation signature in the metadata
ok($meta->generator() eq "OpenOffice::OODoc 1.105 installation test");

exit 0;

