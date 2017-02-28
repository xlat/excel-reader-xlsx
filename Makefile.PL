
use strict;
use warnings;

use 5.008002;

use ExtUtils::MakeMaker 6.30;



my %WriteMakefileArgs = (
  "ABSTRACT" => "Efficient data reader for the Excel XLSX file format.\r",
  "AUTHOR" => "John McNamara <jmcnamara\@cpan.org>",
  "BUILD_REQUIRES" => {},
  "CONFIGURE_REQUIRES" => {
    "ExtUtils::MakeMaker" => "6.30"
  },
  "DISTNAME" => "excel-reader-xlsx",
  "EXE_FILES" => [],
  "LICENSE" => "perl",
  "NAME" => "excel::reader::xlsx",
  "PREREQ_PM" => {
    "XML::Simple::Tiny" => 0,   #https://github.com/xlat/XML-Simple-Tiny
    "Archive::Zip" => 0,
    "Carp" => 0,
    "Exporter" => 0,
    "File::Temp" => 0,
    "OLE::Storage_Lite" => 0,
    "XML::LibXML::Reader" => 0,
    "strict" => 0,
    "version" => 0,
    "warnings" => 0
  },
  "TEST_REQUIRES" => {
    "JSON" => 0,
    "Test::Differences" => 0,
    "Test::More" => 0,
    "lib" => 0
  },
  "VERSION" => "v0.001_010",
  "test" => {
    "TESTS" => "t/package/content_types/*.t t/package/shared_strings/*.t t/regression/*.t t/regression/workbook/*.t t/regression/worksheet/*.t"
  }
);


unless ( eval { ExtUtils::MakeMaker->VERSION(6.63_03) } ) {
  my $tr = delete $WriteMakefileArgs{TEST_REQUIRES};
  my $br = $WriteMakefileArgs{BUILD_REQUIRES};
  for my $mod ( keys %$tr ) {
    if ( exists $br->{$mod} ) {
      $br->{$mod} = $tr->{$mod} if $tr->{$mod} > $br->{$mod};
    }
    else {
      $br->{$mod} = $tr->{$mod};
    }
  }
}

unless ( eval { ExtUtils::MakeMaker->VERSION(6.56) } ) {
  my $br = delete $WriteMakefileArgs{BUILD_REQUIRES};
  my $pp = $WriteMakefileArgs{PREREQ_PM};
  for my $mod ( keys %$br ) {
    if ( exists $pp->{$mod} ) {
      $pp->{$mod} = $br->{$mod} if $br->{$mod} > $pp->{$mod};
    }
    else {
      $pp->{$mod} = $br->{$mod};
    }
  }
}

delete $WriteMakefileArgs{CONFIGURE_REQUIRES}
  unless eval { ExtUtils::MakeMaker->VERSION(6.52) };

WriteMakefile(%WriteMakefileArgs);



