#!/usr/bin/perl -w
use strict;

use Test::More tests => 7;

use Mail::Outlook;

my $outlook = new Mail::Outlook();
isa_ok($outlook,'Mail::Outlook');

my $message = $outlook->create();
isa_ok($message,'Mail::Outlook::Message');

$message->To('you@example.com');
$message->Cc('Them <them@example.com>');
$message->Bcc('Us <us@example.com>; anybody@example.com');
$message->Subject('Blah Blah Blah');
$message->Body('Yadda Yadda Yadda');

is($message->To(),'you@example.com');
is($message->Cc(),'Them <them@example.com>');
is($message->Bcc(),'Us <us@example.com>; anybody@example.com');
is($message->Subject(),'Blah Blah Blah');
is($message->Body(),'Yadda Yadda Yadda');
