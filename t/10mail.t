use Test::More tests => 7;

SKIP: {
	eval "use Win32::OLE::Const 'Microsoft Outlook'";
	skip "Microsoft Outlook doesn't appear to be installed\n", 7	if($@);

	eval "use Mail::Outlook";
	skip "Unable to make a connection to Microsoft Outlook\n", 7	if($@);

	my $outlook = Mail::Outlook->new();
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
}
