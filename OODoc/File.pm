#-----------------------------------------------------------------------------
#
#	$Id : File.pm 1.103 2004-03-07 JMG$
#
#	Initial developer: Jean-Marie Gouarne
#	Copyright 2004 by Genicorp, S.A. (www.genicorp.com)
#	Licensing conditions:
#		- Licence Publique Generale Genicorp v1.0
#		- GNU Lesser General Public License v2.1
#	Contact: oodoc@genicorp.com
#
#-----------------------------------------------------------------------------

package	OpenOffice::OODoc::File;
use	5.006_001;
our	$VERSION	= 1.103;
use	Archive::Zip	1.05	qw ( :DEFAULT :CONSTANTS :ERROR_CODES );

#-----------------------------------------------------------------------------
# some defaults

our	$DEFAULT_COMPRESSION_METHOD	= COMPRESSION_DEFLATED;
our	$DEFAULT_COMPRESSION_LEVEL	= COMPRESSION_LEVEL_BEST_COMPRESSION;
our	$DEFAULT_EXPORT_PATH		= './';

#-----------------------------------------------------------------------------
# control & conversion of XML component names of the OO file

sub	CtrlMemberName
	{
	my $self	= shift;
	my $member	= shift;

	my $m = lc $member;
	foreach my $n ('content', 'meta', 'styles', 'settings')
		{
		if ($m eq $n)
			{
			$member = $n . '.xml';
			last;
			}
		}

	foreach my $m ( @{ $self->{'members'} } )
		{
		return $member if ($member eq $m);
		}

	return	undef;
	}

#-----------------------------------------------------------------------------
# error handler for low-level zipfile errors (to be customized if needed)

sub	ZipError
	{
	warn "[" . __PACKAGE__ . "] Zip Alert\n";
	}

#-----------------------------------------------------------------------------
# unique temporary file name generation

sub	new_temp_file_name
	{
	my $self	= shift;

	return Archive::Zip::tempFileName($self->{'work_dir'});
	}

#-----------------------------------------------------------------------------
# temporary data storage

sub	store_temp_file
	{
	my $self	= shift;
	my $data	= shift;

	my $tmpfile	= $self->new_temp_file_name;
	unless (open FH, '>', $tmpfile)
		{
		warn	"[" . __PACKAGE__ . "::store_temp_file] "	.
			"Unable to create temporary file $tmpfile\n";
		return undef;
		}
	unless (print FH $data)
		{
		warn	"[" . __PACKAGE__ . "::store_temp_file] "	.
			"Write error in temporary file $tmpfile\n";
		return undef;
		}
	unless (close FH)
		{
		warn	"[" . __PACKAGE__ . "::store_temp_file] "	.
			"Unknown error in temporary file $tmpfile\n";
		return undef;
		}
	push @{$self->{'temporary_files'}}, $tmpfile;
	return $tmpfile;
	}

#-----------------------------------------------------------------------------
# temporary member extraction

sub	extract_temp_file
	{
	my $self	= shift;
	my $member	= shift;

	my $m	= ref $member	?
			$member		:
			$self->{'archive'}->memberNamed($member);

	my $tmpfile	= $self->new_temp_file_name;

	my $result = $m->extractToFileNamed($tmpfile);
	if ($result == AZ_OK)
		{
		push @{$self->{'temporary_files'}}, $tmpfile;
		return $tmpfile;
		}
	else
		{
		return undef;
		}
	}

#-----------------------------------------------------------------------------
# temporary  storage cleanup
# returns the number of deleted files and clears the list of temp files

sub	remove_temp_files
	{
	my $self	= shift;
	my $count	= 0;

	while (@{$self->{'temporary_files'}})
		{
		my $tmpfile	= shift @{$self->{'temporary_files'}};
		my $r		= unlink $tmpfile;
		unless ($r > 0)
			{
			warn
				"[" . __PACKAGE__ . "::remove_temp_files] " .
				"Temporary file $tmpfile can't be removed\n";
			}
		else
			{
			$count++;
			}
		}
	return $count;
	}

#-----------------------------------------------------------------------------
# constructor; requires an existing regular OO file

sub	new
	{
	my $caller	= shift;
	my $class	= ref($caller) || $caller;
	my $sourcefile	= shift;
	my $self	= 	{
				'linked'		=> [],
				'work_dir'		=> '/tmp',
				'temporary_files'	=> [],
				'raw_members'		=> [],
				@_
				};
	
	$self->{'source_file'}	= $sourcefile;

	unless	($sourcefile)
		{
		warn "[" . __PACKAGE__ . "::new] No source file name\n";
		return undef;
		}
	
	unless	( -e $sourcefile && -f $sourcefile && -r $sourcefile )
		{
		warn	"[" . __PACKAGE__ .
			"::new] File $sourcefile unavailable\n";
		return undef;
		}
	
	Archive::Zip::setErrorHandler ( \&ZipError );

	$self->{'archive'}	= Archive::Zip->new;

	if ($self->{'source_file'})
		{
		if ($self->{'archive'}->read($self->{'source_file'}) != AZ_OK)
			{
			warn "[" . __PACKAGE__ . "::new] Read error\n";
			$self->{'archive'}	= undef;
			return undef;
			}
		$self->{'members'}	= [ $self->{'archive'}->memberNames ];
		}

	return bless $self, $class;
	}

#-----------------------------------------------------------------------------
# individual zip XML member extraction/uncompression

sub	extract
	{
	my $self	= shift;
	my $member	= $self->CtrlMemberName(shift);

	unless ($member)
		{
		warn "[" . __PACKAGE__ . "::extract] Unknown member\n";
		return undef;
		}

	unless ($self->{'archive'})
		{
		warn "[" . __PACKAGE__ . "::extract] No archive\n";
		return undef;
		}

	return	$self->{'archive'}->contents($member);
	}

#-----------------------------------------------------------------------------
# individual zip member raw export (see Archive::Zip::extractMember)

sub	raw_export
	{
	my $self	= shift;

	unless ($self->{'archive'})
		{
		warn "[" . __PACKAGE__ . "::raw_export] No archive\n";
		return undef;
		}

	my $source	= shift;
	my $target	= shift;
	if (defined $target)
		{
		unless ($target =~ /\//)
			{
			$target = $DEFAULT_EXPORT_PATH . $target;
			}
		unshift @_, $target;
		}
	unshift @_, $source;

	if ($self->{'archive'}->extractMember(@_) == AZ_OK)
		{
		return	$target	? $target : $source;
		}
	else
		{
		warn "[" . __PACKAGE__ . "::raw_export] File output error\n";
		return undef;
		}
	}

#-----------------------------------------------------------------------------
# individual zip member raw import
# file to be imported is only registered here; real import by save()

sub	raw_import
	{
	my $self	= shift;
	my $membername	= shift;
	my $filename	= shift;
	$filename	= $membername	unless $filename;
	my %new_member	= ('file' => $filename, 'member' => $membername);

	push @{$self->{'raw_members'}}, \%new_member;
	return %new_member;
	}

#-----------------------------------------------------------------------------
# individual zip member removing (real deletion committed by save)
# WARNING: removing a member doesn't automatically update "manifest.xml"

sub	raw_delete
	{
	my $self	= shift;
	my $member	= shift;

	my $mbcount	= scalar @{$self->{'members'}};
	for (my $i = 0 ; $i < $mbcount ; $i++)
		{
		if ($self->{'members'}[$i] eq $member)
			{
			splice(@{$self->{'members'}}, $i, 1);
			return 1;
			}
		}
	return undef;
	}

#-----------------------------------------------------------------------------
# archive list

sub	getMemberList
	{
	my $self	= shift;

	return @{$self->{'members'}};
	}

#-----------------------------------------------------------------------------
# connects the current OODoc::File instance to a client OODoc::XPath object
# and extracts the corresponding XML member (to be transparently invoked
# by the constructor of OODoc::XPath when activated with a 'file' parameter)

sub	link
	{
	my $self	= shift;
	my $ooobject	= shift;

	push @{$self->{'linked'}}, $ooobject;
	return $self->extract($ooobject->{'member'});
	}

#-----------------------------------------------------------------------------
# copy an individual member from the current OODoc::File instance($self)
# to an external Archive::Zip object ($archive), using a temporary flat file

sub	copyMember
	{
	my $self	= shift;
	my $archive	= shift;
	my $member	= shift;

	my $m		= $self->{'archive'}->memberNamed($member);
	unless ($m)
		{
		warn	"[" . __PACKAGE__ .
			"::copyMember] Unknown source member\n";
		return undef;
		}
	my $tmpfile	= $self->extract_temp_file($m);
	unless ($tmpfile)
		{
		warn	"[" . __PACKAGE__ .
			"::copyMember] File extraction error\n";
		return undef;
		}
		
	$m	= $archive->addFile($tmpfile, $member);
	unless ($member eq 'meta.xml')
		{
		$m->desiredCompressionMethod
				($DEFAULT_COMPRESSION_METHOD);
		$m->desiredCompressionLevel
				($DEFAULT_COMPRESSION_LEVEL);
		}
	else
		{
		$m->desiredCompressionMethod(COMPRESSION_STORED);
		}
	}

#-----------------------------------------------------------------------------
# inserts $data as a new member in an external Archive::Zip object

sub	addNewMember
	{
	my $self			= shift;
	my ($archive, $member, $data)	= @_;

	unless ($archive && $member && $data)
		{
		warn	"[" . __PACKAGE__ .
			"::addNewMember] Missing argument(s)\n";
		return undef;
		}

				# temporary file creation --------------------

	my $tmpfile = $self->store_temp_file($data);
	unless ($tmpfile)
		{
		warn	"[" . __PACKAGE__ . "::addNewMember] "	.
			"Temporary file error\n";
		return undef;
		}
				
				# member insertion/compression ---------------

	my $m = $archive->addFile($tmpfile, $member);
	unless ($member eq 'meta.xml')
		{
		$m->desiredCompressionMethod
				($DEFAULT_COMPRESSION_METHOD);
		$m->desiredCompressionLevel
				($DEFAULT_COMPRESSION_LEVEL);
		}
	else
		{
		$m->desiredCompressionMethod(COMPRESSION_STORED);
		}
	return $m;
	}

#-----------------------------------------------------------------------------
# creates a new OO file, copying unchanged members & updating
# modified ones (by default, the new OO file replaces the old one)

sub	save
	{
	my $self		= shift;
	my $targetfile		= shift;

	my %newmembers		= ();
	foreach my $nm (@{$self->{'linked'}})
		{
		$newmembers{$nm->{'member'}} = $nm->getXMLContent;
		}

	my $outfile	= undef;
	my $tmpfile	= undef;

				# target file check --------------------------

	$targetfile	= $self->{'source_file'}	unless $targetfile;
	if ( -e $targetfile )
		{
		unless ( -f $targetfile )
			{
			warn 	"[" . __PACKAGE__ . "::save] "	.
				"$targetfile is not a regular file\n";
			return undef;
			}
		unless ( -w $targetfile )
			{
			warn 	"[" . __PACKAGE__ . "::save "	.
				"$targetfile is read only\n";
			return undef;
			}
		}
		
				# output to temporary file if target eq source

	if ($targetfile eq $self->{'source_file'})
		{
		$outfile = $self->new_temp_file_name;
		}
	else
		{
		$outfile = $targetfile;
		}

				# discriminate replaced/added members --------

	my %replacedmembertable	= ();
	my @addedmemberlist	= ();
	foreach my $nmn (keys %newmembers)
		{
		my $tmn	= $self->CtrlMemberName($nmn);
		if ($tmn)
			{
			$replacedmembertable{$tmn}	= $nmn;	
			}
		else
			{
			push @addedmemberlist, $nmn;
			}
		}
		
				# target archive operation -------------------

	my $archive	= Archive::Zip->new;
	foreach my $oldmember (@{$self->{'members'}})
		{
		my $k	= $replacedmembertable{$oldmember};
		unless ($k)			# (unchanged member)
			{
			$self->copyMember($archive, $oldmember);
			}
		else				# (replaced member)
			{
			$self->addNewMember
				($archive, $oldmember, $newmembers{$k});
			}
		}
	foreach my $name (@addedmemberlist)	# (added member)
		{
		$self->addNewMember($archive, $name, $newmembers{$name});
		}
	
	foreach my $raw_member (@{$self->{'raw_members'}}) # optional raw data
		{
		$archive->removeMember($raw_member->{'member'});
		my $m = $archive->addFile
			($raw_member->{'file'}, $raw_member->{'member'});
		$m->desiredCompressionMethod
			($DEFAULT_COMPRESSION_METHOD);
		$m->desiredCompressionLevel
			($DEFAULT_COMPRESSION_LEVEL);
		}

	my $status = $archive->writeToFileNamed($outfile);

				# post write control & cleanup ---------------
	if ($status == AZ_OK)
		{
		unless ($outfile eq $targetfile)
			{
			require File::Copy;

			unlink $targetfile;
			File::Copy::move($outfile, $targetfile);
			}
		$self->remove_temp_files;
		return 1;
		}
	else
		{
		warn "[" . __PACKAGE__ . "::save] Archive write error\n";
		return undef;
		}
	}

#-----------------------------------------------------------------------------
1;
