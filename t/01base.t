use Test::More tests => 3;

SKIP: {
	eval "use Win32::OLE::Const 'Microsoft Outlook'";
	skip "Microsoft Outlook doesn't appear to be installed\n", 3	if($@);

	use_ok( 'Mail::Outlook' );
	use_ok( 'Mail::Outlook::Folder' );
	use_ok( 'Mail::Outlook::Message' );
}

