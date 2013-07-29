#!/usr/bin/perl -w
use strict;
no strict 'subs';

use Test::More tests => 12;

use lib 't/testlib';

my $tests = 12;

eval {

SKIP: {
	eval "use Typelibs";
	skip "Microsoft Outlook doesn't appear to be installed\n", $tests	if($@);

	my $vers = Typelibs::ExistsTypeLib('Microsoft Outlook');
	skip "Microsoft Outlook doesn't appear to be installed\n", $tests	unless($vers);

	eval "use Mail::Outlook";
	skip "Unable to make a connection to Microsoft Outlook\n", $tests	if($@);

	my $outlook = Mail::Outlook->new('Inbox');
	my $folder = $outlook->folder();
	isa_ok($folder,'Mail::Outlook::Folder');

	my $message = $folder->first;
	isa_ok($message,'Mail::Outlook::Message');

    my $name = $message->From();
    skip "Access to Microsoft Outlook has been declined", ($tests - 2)  unless($name);
	ok($name);

	$message = $folder->next;
	isa_ok($message,'Mail::Outlook::Message');
	ok($message->From());

	$folder = $outlook->folder('Inbox');
	$message = $folder->last;
	isa_ok($message,'Mail::Outlook::Message');
	ok($message->From());

	$message = $folder->previous;
	isa_ok($message,'Mail::Outlook::Message');
	ok($message->From());

	$folder = $outlook->folder('Inbox/ANameThatShouldNotExist');
	is($folder,undef);
	$folder = $outlook->folder('ANameThatShouldNotExist');
	is($folder,undef);

	eval "use Win32::OLE::Const 'Microsoft Outlook'";
	skip "Unable to make a connection to Microsoft Outlook\n", 1	if($@);

	$folder = $outlook->folder(olFolderInbox);
	isa_ok($folder,'Mail::Outlook::Folder');
}

};

if($@ =~ /Network problems/) {
	skip "Microsoft Outlook cannot connect to the server.\n", $tests;
	exit;
}
