#-----------------------------------------------------------------------------
# 01write.t	OpenOffice::OODoc Installation test		(c) GENICORP
#-----------------------------------------------------------------------------

use Test;
BEGIN	{ plan tests => 20 }

use OpenOffice::OODoc	1.309;
ok($OpenOffice::OODoc::VERSION >= 1.309);

#-----------------------------------------------------------------------------

my $generator	=	"OpenOffice::OODoc " . $OpenOffice::OODoc::VERSION .
			" installation test";
my $testfile	=	"ootest.sxw";
my $class	=	"text";
my $imagefile	=	"OODoc/data/scarab.png";
my $test_date	=	ooLocaltime();

# Creating an empty new OpenOffice.org file with the default template
unlink $testfile;
my $archive = ooFile($testfile, create => $class);
unless ($archive)
	{
	ok(0); # Unable to create the test file
	exit;
	}
else
	{
	ok(1); # Test file created
	}

#-----------------------------------------------------------------------------

my $notice	= 
"This OpenOffice.org document has been generated by the OpenOffice::OODoc " .
"installation test. If you can read this paragraph in blue letters with " .
"a yellow background, and if you can see a green scarab at the top of the " .
"page, your installation is probably OK.";

my $title	= "OpenOffice::OODoc test document";
my $description	= "Generated by $generator";

# Opening the content using OpenOffice::OODoc::Document
my $doc	= ooDocument(archive => $archive, member => 'content')
	or die "# Unable to find a regular OpenOffice.org content\n";
ok($doc); # Document open and parsed

my $styles = ooDocument(archive => $archive, member => 'styles')
	or warn "# Unable to get the styles\n";
ok($styles); # Styles open and parsed

# Creating a graphic style
ok($styles->createImageStyle('Logo'));
# Inserting an image in the document
ok	(
	$doc->createImageElement
		(
		"Logo",
		style		=> 'Logo',
		page		=> 1,
		position	=> '2cm, 2cm',
		size		=> '2cm, 2.824cm',
		import		=> $imagefile
		)
	);

# Appending a paragraph
ok	(
	$doc->appendParagraph( text => "File creation date : " . localtime() )
	);
# Appending a level 1 header
ok	(
	$doc->appendHeader( text => "Congratulations !", level => "1" )
	);
# Creating a coloured paragraph style (blue foreground, yellow background)
ok	(
	$styles->createStyle
		(
		"Colour",
		family		=> 'paragraph',
		parent		=> 'Standard',
		properties	=>
			{
			'fo:color'		=> rgb2oo(0,0,128),
			'fo:background-color'	=> rgb2oo("yellow")
			}
		)
	);
# Appending another paragraph using the new style
ok	(
	$doc->appendParagraph( text => $notice, style => "Colour" )
	);
# Appending another level 1 header
ok	(
	$doc->appendHeader( text => "Your environment :", level => "1" )
	);
# Appending an item list with 5 elements
$doc->setText
	(
	$doc->appendItemList,
	"Platform : $^O",
	"Perl version : $]",
	"Archive::Zip version : $Archive::Zip::VERSION",
	"XML::Twig version : $XML::Twig::VERSION",
	"OpenOffice::OODoc version : $OpenOffice::OODoc::VERSION"
	);
my $list = $doc->getUnorderedList(0);
my $count = scalar $doc->selectElements($list, 'text:list-item');
ok($count == 5);

# Appending another level 1 header
ok	(
	$doc->appendHeader( text => "Your OpenOffice::OODoc choices :", level => "1" )
	);
# Appending an item list with the installation parameters
$doc->setText
	(
	$doc->appendItemList,
	"Local character set : $OpenOffice::OODoc::XPath::LOCAL_CHARSET",
	"Working directory : $OpenOffice::OODoc::File::WORKING_DIRECTORY"
	);

# Opening the metadata of the document
my $meta = ooMeta(archive => $archive)
	or die "# Unable to find regular OpenOffice.org metadata\n";
ok($meta);
# Checking the title of the document
ok($meta->title($title));
ok($meta->description($description));
ok($meta->generator($generator));
ok($meta->creation_date($test_date));
ok($meta->date($test_date));

# Saving the $testfile file
ok($archive->save);

exit 0;
