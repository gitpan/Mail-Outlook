#!/usr/bin/perl -w
use strict;

#########################

use Test::More tests => 3;

BEGIN {
	use_ok( 'Mail::Outlook' );
	use_ok( 'Mail::Outlook::Folder' );
	use_ok( 'Mail::Outlook::Message' );
}

#########################

