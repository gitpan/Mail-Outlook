use ExtUtils::MakeMaker;
WriteMakefile(
    'NAME'		=> 'Mail::Outlook',
    'VERSION_FROM'	=> 'lib/Mail/Outlook.pm',
    'PREREQ_PM'		=> {
		Win32::OLE => 0,
		Win32::OLE::Const => 0,
	},
    ($] >= 5.005 ?
      (ABSTRACT_FROM => 'lib/Mail/Outlook.pm',
       AUTHOR     => 'Barbie <barbie@cpan.org>') : ()),
);
