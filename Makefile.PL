use 5.008;
use ExtUtils::MakeMaker;

WriteMakefile(
    'NAME'		=> 'OpenOffice::OODoc',
    'VERSION_FROM'	=> 'OODoc.pm',
    'AUTHOR'		=> 'Jean-Marie Gouarne, Genicorp S.A.',
    'PREREQ_PM'		=>
    		{
		'XML::XPath'		=> '1.13',
		'Archive::Zip'		=> '1.05'
		}
);
