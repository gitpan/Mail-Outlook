#!/usr/bin/perl -w
use strict;

use Test::More tests => 6;

use Mail::Outlook;

my %hash = (
	To		=> 'you@example.com',
	Cc		=> 'Them <them@example.com>',
	Bcc		=> 'Us <us@example.com>; anybody@example.com',
	Subject	=> 'Blah Blah Blah',
	Body	=> 'Yadda Yadda Yadda',
  );

my $outlook = new Mail::Outlook();
my $message = $outlook->create(%hash);
isa_ok($message,'Mail::Outlook::Message');

is($message->To(),'you@example.com');
is($message->Cc(),'Them <them@example.com>');
is($message->Bcc(),'Us <us@example.com>; anybody@example.com');
is($message->Subject(),'Blah Blah Blah');
is($message->Body(),'Yadda Yadda Yadda');

