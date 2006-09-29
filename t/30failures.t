use Test::More tests => 3;

SKIP: {
	my $tests = 3;

	eval "use Test::MockObject";
	skip "Test::MockObject doesn't appear to be installed\n", $tests	if($@);

	my $mock = Test::MockObject->new();
	$mock->fake_module( 'Win32::OLE' );

	eval "use Mail::Outlook";
	skip "Unable to make a connection to Microsoft Outlook\n",$tests	if($@);

	$mock->mock( 'GetActiveObject', sub { die "Forced Failure" } );
	my $outlook = Mail::Outlook->new();
	is($outlook,undef);

	$mock->mock( 'GetActiveObject', sub { return undef } );
	$outlook = Mail::Outlook->new();
	is($outlook,undef);

	$mock->mock( 'GetNameSpace', sub { return undef } );
	$outlook = Mail::Outlook->new();
	is($outlook,undef);
}
