use 5.008;
use Getopt::Long;
use Encode;
use ExtUtils::MakeMaker;

#---

GetOptions
	(
	'config!'	=> \(my $config = 1),
	'prompt!'	=> \(my $prompt = 1),
	'encoding=s'	=> \(my $encoding = 'iso-8859-1'),
	'colormap=s'	=> \(my $colormap = ''),
	'workdir=s'	=> \(my $workdir = '.')
	);

#---

if ($config)
	{
	my $config_file = 'OODoc/config.xml';
	print	"\nNow you will be prompted for some configuration options.\n"	.
		"These options are intended to define default values only.\n"	.
		"Each option may be overridden by the applications.\n"		.
		"You can change these default options later; to do so, you\n"	.
		"can either replay this installation procedure or manually\n"	.
		"edit the <installation path>/$config_file file.\n"		.
		"(See INSTALL for details)\n"
		if ($prompt);
	while ($prompt)
		{
		my $answer;
		print "\nPresent configuration:\n";
		print "- Your local character set is now [$encoding]\n";
		unless (Encode::find_encoding($encoding))
			{
			warn	"\tWARNING: Unsupported character set.\n";
			}
		unless ($colormap)
			{
			print "- There is no colour map file\n";
			}
		else
			{
			print "- Your colour map file is now [$colormap]\n";
			unless (-f $colormap)
				{
				warn "\tWARNING: This file doesn't exist.\n";
				}
			}
		print "- Your working directory for temporary files is now [$workdir]\n";
		unless (-d $workdir)
			{
			warn	"\tWARNING: This directory doesn't exist.\n";
			}
		$answer = lc (prompt("Is that OK (y/n) ?", "y"));
		$prompt = '' unless $answer ne "y";
		last unless $prompt;
		$encoding	= prompt
			(
			"What is your preferred local character set ?",
			$encoding
			);
		$colormap	= prompt
			(
			"What is the full path of your RGB colour map file (optional) ?",
			$colormap
			);
		$workdir	= prompt
			(
			"What is your working directory for temporary files ?",
			$workdir
			);
		$encoding = '' unless $encoding gt ' ';
		$colormap = '' unless $colormap gt ' ';
		$workdir  = '' unless $workdir gt ' ';
		}
	
	Encode::from_to($workdir, $encoding, 'utf8');

#---

	my @lt = localtime();
	my $install_date = sprintf
		(
		"%04d-%02d-%02dT%02d:%02d:%02d",
		$lt[5] + 1900, $lt[4] + 1, $lt[3], $lt[2], $lt[1], $lt[0]
		);

#---

	open CF, ">", $config_file;
	print CF '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
	print CF "<config>\n";
	print CF "\t<comment>OpenOffice::OODoc local configuration file</comment>\n";
	print CF "\t<OpenOffice-OODoc>\n";
	print CF "\t\t<XPath-LOCAL_CHARSET>$encoding</XPath-LOCAL_CHARSET>\n";
	print CF "\t\t<File-WORKING_DIRECTORY>$workdir</File-WORKING_DIRECTORY>\n";
	print CF "\t\t<Styles-COLORMAP>$colormap</Styles-COLORMAP>\n";
	print CF "\t\t<INSTALLATION_DATE>$install_date</INSTALLATION_DATE>\n";
	print CF "\t</OpenOffice-OODoc>\n";
	print CF "</config>";
	close CF;
	}

#---

WriteMakefile
	(
    	'NAME'		=> 'OpenOffice::OODoc',
    	'VERSION_FROM'	=> 'OODoc.pm',
    	'ABSTRACT_FROM'	=> 'OODoc.pod',
    	'AUTHOR'	=> 'Jean-Marie Gouarne, Genicorp S.A.',
    	'PREREQ_PM'	=>
    		{
		'XML::Twig'		=> '3.15',
		'Archive::Zip'		=> 0,
		'File::Temp'		=> '0.12',
		'Time::Local'		=> '1.07'
		},
	'EXE_FILES'	=>
		[
		'examples/oohighlight',
		'examples/oofilesearch',
		'examples/text2ooo',
		'examples/text2table',
		'examples/oo2pod',
		'examples/oobuild',
		'examples/oometadoc'
		]
	);

#---
