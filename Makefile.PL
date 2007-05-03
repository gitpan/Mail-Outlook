use ExtUtils::MakeMaker;
WriteMakefile(
    'NAME'		    => 'Mail::Outlook',
    'VERSION_FROM'	=> 'lib/Mail/Outlook.pm',
    'PREREQ_PM'		=> {
        File::Basename      => 0,
		Win32::OLE          => 0,
		Win32::OLE::Const   => 0,
	},
	NO_META			=> 1,
    ($] >= 5.005 ?
      (ABSTRACT_FROM => 'lib/Mail/Outlook.pm',
       AUTHOR       => 'Barbie <barbie@cpan.org>') : ()),
);

print <<HERE;

  Please note that Microsoft Outlook is required by this distribution.


  PLEASE NOTE:
  1) You will see several Win32::OLE errors during testing. This is expected 
     as we should try and test failure as well as success :)
  2) 1 message is sent during testing. If Outlook requests your acceptance,
     you may decline, as the testing will acknowledge this.
  3) Part of the testing uses your Inbox to validate whether the distribution
     can read mail. If you are uncomfortable with this, please decline any
     Outlook security pop up and those tests will be skipped.

HERE
