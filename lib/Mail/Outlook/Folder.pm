package Mail::Outlook::Folder;

use warnings;
use strict;

use vars qw($VERSION);
$VERSION = '0.03';

#----------------------------------------------------------------------------

=head1 NAME

Mail::Outlook::Folder - extension to handle Microsoft (R) Outlook (R) mail folders.

=head1 SYNOPSIS

See Mail::Outlook, as this is not meant to be used as a standalone module.

=head1 DESCRIPTION

Handles the Folder interaction with the Outlook API.

=cut

#----------------------------------------------------------------------------

#############################################################################
#Library Modules															#
#############################################################################

use Win32::OLE;
use Win32::OLE::Const 'Microsoft Outlook';

use Mail::Outlook::Message;

#############################################################################
#Variables
#############################################################################

my %foldernames = (
	'Inbox'			=> olFolderInbox,
	'Outbox'		=> olFolderOutbox,
	'Sent Items'	=> olFolderSentMail,
);

#----------------------------------------------------------------------------

#############################################################################
#Interface Functions														#
#############################################################################

=head1 METHODS

=over 4

=item new()

Create a new Outlook mail object. Returns the object on success or undef on
failure. To see the last error use 'Win32::OLE->LastError();'.

=cut

sub new {
	my ($self, $outlook, $foldername) = @_;
	my ($mailbox,$folder);

	# mailbox and path
	if ($foldername =~ m!/!) {
		my ($box,$name) = ($foldername =~ m!(.*?)/(.*)!);
		if($foldernames{$box}) {
			$mailbox = $outlook->{namespace}->GetDefaultFolder($foldernames{$box})
				or return undef;
			$folder = $mailbox->Folders($name)
				or return undef;
		} else {
			return undef;
		}

	# mailbox only
	} elsif($foldernames{$foldername}) {
		$mailbox = $outlook->{namespace}->GetDefaultFolder($foldernames{$foldername})
			or return undef;

	# mailbox constant only
	} elsif(defined $foldername) {
		$mailbox = $outlook->{namespace}->GetDefaultFolder($foldername)
			or return undef;
		$foldername = 'Not Known';
		
	# well if you don't know, neither do i!!!
	} else {
		return undef;
	}


	# create an attributes hash
	my $atts = {
		'outlook'		=> $outlook,
		'foldername'	=> $foldername,
		'objfolder'		=> $folder || $mailbox || undef,
		'items'			=> undef,
	};

	# prime the mail items collection
	$atts->{items} = $atts->{objfolder}->Items()	or return undef;


	# create the object
	bless $atts, $self;
	return $atts;
}

sub DESTROY {}

=item first()

Gets the first message object in the current folder. Returns undef if no messages.

=cut

sub first {
	my $self = shift;
	return $self->{items}->GetFirst();
}

=item last()

Gets the last message object in the current folder. Returns undef if no messages.

=cut

sub last {
	my $self = shift;
	return $self->{items}->GetLast();
}

=item next()

Gets the next message object in the current folder. Returns undef if no more 
messages. Must be called after a first() or last() has been intiated.

=cut

sub next {
	my $self = shift;
	return $self->{items}->GetNext();
}

=item previous()

Gets the previous message object in the current folder. Returns undef if no 
more messages. Must be called after a first() or last() has been intiated.

=cut

sub previous {
	my $self = shift;
	return $self->{items}->GetPrevious();
}

1;


__END__

#----------------------------------------------------------------------------

=back

=head1 BUGS, PATCHES & FIXES

There are no known bugs at the time of this release. However, if you spot a
bug or are experiencing difficulties that are not explained within the POD
documentation, please send an email to barbie@cpan.org or submit a bug to the
RT system (http://rt.cpan.org/). However, it would help greatly if you are 
able to pinpoint problems or even supply a patch. 

If you intend to supply a patch, please visit the following URL (and 
associated pages) to ensure you are using the correct objects and methods.

http://msdn.microsoft.com/library/default.asp?url=/library/en-us/off2000/html/olobjApplication.asp

This article contains some interesting background into creating mail
messages via Outlook, although it is VB-centric.

http://www.exchangeadmin.com/Articles/Index.cfm?ArticleID=4657

Fixes are dependant upon their severity and my availablity. Should a fix not
be forthcoming, please feel free to (politely) remind me.

=head1 FUTURE ENHANCEMENTS

A couple of items that I'd like to get working.

* X-Header support
* Send without the popups (Outlook Redemption looks possible)

=head1 NOTES

This module is intended to be used on Win32 platforms only, with Microsoft (R) 
Outlook (R) installed.

  Microsoft and Outlook are registered trademarks and the copyright 1995-2003
  of Microsoft Corporation. All rights reserved.

=head1 SEE ALSO

  Win32::OLE
  Win32::OLE::Const

=head1 AUTHOR

Barbie, C< <<barbie@cpan.org>> >
for Miss Barbell Productions, L<http://www.missbarbell.co.uk>

Birmingham Perl Mongers, L<http://birmingham.pm.org/>

=head1 COPYRIGHT AND LICENSE

  Copyright (C) 2003-2005 Barbie for Miss Barbell Productions
  All Rights Reserved.

  This module is free software; you can redistribute it and/or 
  modify it under the same terms as Perl itself.

=cut
