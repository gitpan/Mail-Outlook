use Test::More tests => 6;
	
SKIP: {
	eval "use Win32::OLE::Const 'Microsoft Outlook'";
	skip "Microsoft Outlook doesn't appear to be installed\n", 6	if($@);

	eval "use Mail::Outlook";
	skip "Unable to make a connection to Microsoft Outlook\n", 6	if($@);

	my %hash = (
		To		=> 'you@example.com',
		Cc		=> 'Them <them@example.com>',
		Bcc		=> 'Us <us@example.com>; anybody@example.com',
		Subject	=> 'Blah Blah Blah',
		Body	=> 'Yadda Yadda Yadda',
  	);

	my $outlook = Mail::Outlook->new();
	my $message = $outlook->create(%hash);
	isa_ok($message,'Mail::Outlook::Message');

	is($message->To(),'you@example.com');
	is($message->Cc(),'Them <them@example.com>');
	is($message->Bcc(),'Us <us@example.com>; anybody@example.com');
	is($message->Subject(),'Blah Blah Blah');
	is($message->Body(),'Yadda Yadda Yadda');
}
