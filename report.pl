#!/home/scopych/perl5/perlbrew/perls/perl-5.36.0/bin/perl

use v5.36;
use warnings;
use strict;
use feature qw(say);
use Encode qw(encode decode);
#use utf8;
use Term::ANSIColor;
use Excel::ValueReader::XLSX;
use Data::Dump qw(dump);
use Data::Printer;

my $filename = "$ARGV[0]";

#open(IN, "<$in_file") or warn "Cannot open $in_file\n";
#my @text = <IN>;
my @ru_card_fields = qw(имя рождение крещение назначение пол надежда служ_год месяци);
my @ru_months = qw(сентябрь октябрь ноябрь декабрь январь февраль март апрель май июнь июль август);
my @ru_assignment = ('старейшина', 'служебный помощьник', 'пионер', 'подсобный пионер', 'возвещатель', 'некрещенный');
my @cards;

my %values =
(	houers => "",
	studys => "",
	notes  => "",
);

my %months =
(	september => "",
	october   => "",
	november  => "",
	december  => "",
	january   => "",
	february  => "",
	march     => "",
	april     => "",
	may       => "",
	june      => "",
	july      => "",
	august    => "",
);

my %card =
(	name            => "",
	birth           => "",
	baptism         => "",
	assignment      => "",
	sex             => "",
	hope            => "",
	year_of_service => "",
	month           => \%months,
);

my $reader = Excel::ValueReader::XLSX->new(xlsx  => $filename);
my @sheets = $reader->sheet_names;
my $last_sheet = $sheets[-1];
my $grid = $reader->values($last_sheet);


sub fill_card
{	
	$card{name} = encode('UTF-8', $grid->[5][1]);
	$values{houers} = $grid->[5][2];
	$values{studys} = $grid->[5][3];
	$values{notes}  = $grid->[5][4];
	$months{may} = \%values;
	$card{month} = \%months;
	say dump(%card);
}

fill_card();
p %card;
			
=begin

fill_cards();
#find_grp_number();

my $c = -1;
while ($c != 0)         # main loop
{ print_menu();
  print color('red');
  print ">>>_";
  print color('reset');
  $c = <STDIN>; chomp($c);
  if ($c == 1)
  { print_grp_rep();}
  if ($c == 2)
  { print color('red');
    print "Введи имя или часть имени_";
    print color('reset');
    print_by_name();
  }
  if ($c == 3)
  { print_by_assignment();}
  if ($c == 4)
  { sums();}
  if ($c == 5)
  { sums()}
  if ($c == 6)
  { num_publishers();}
  if ($c == 0)
  { exit;}
}

#close(IN);

sub print_menu
{ print color('blue'); 
  say "      МЕНЮ\n";
  say "0 Выйти\n";
  say "1 Показать весь отчет\n";
  say "2 Показать отчет по имени\n";
  say "3 Показать отчет по назначению\n";
  say "5 Общие цифры\n";
  say "6 Количество возвещателей\n";
  print color('reset');

  return;
}

sub fill_cards 
{ 
  while (<>) 
  { chomp;
     @card{@fields} = ;
     push @cards, \%card; 
    
  }
}

sub print_grp_rep
{ print "Отчет группы № \n";
  foreach (@grp_rep)
  { my $rec = $_;
    print_fields_of_hash($rec);
  }
  sums();

  return;
}

sub print_fields_of_hash
{ my $rec = $_;
  print "\n";
  print "      №:           $rec->{'№'}\n";
  print "      Имя:         $rec->{'имя'}\n"; 
  print "      Часы:        $rec->{'часы'}\n";
  print "      Изучения:    $rec->{'изучения'}\n";
  print "      Примечания:  $rec->{'примечания'}\n";
  print "      Назначение:  $rec->{'назначение'}\n";
  print "      Служил(а)$rec->{'служил(а)'}\n";
  print "\n";

  return;
}

sub print_by_name
{ my $name = <STDIN>; chomp($name);
  my $rec;
  my $isfound = 0;
  foreach (@grp_rep)#$rec (@grp_rep)
  { if ($_->{'имя'} =~ m/$name/i)
    { $rec = $_;
      print_fields_of_hash($rec);
      $isfound = 'true';
    }
  }
  if (!$isfound)
  { print color('yellow');
    print "         Не найдено\n";
    print color('reset');
  }

  return;
}

sub sums
{ my $rec;
  my $field;
  my @fields = qw(часы изучения служил(а));
  my $sum = 0;
  my ($assignment) = @_;

  print "          Общие цифры\n";

  if (defined($assignment))
  { foreach $field (@fields)
    { foreach $rec (@grp_rep)
      { if ($rec->{'назначение'} eq $assignment)
        { $sum += $rec->{$field};}
      }
      printf "%15s: %5d\n", $field, $sum;
      $sum = 0;
    }
    print "\n";
  }

  else
  { foreach $field (@fields)
    { foreach $rec (@grp_rep)
      { $sum += $rec->{$field};}
      printf "%15s: %5d\n", $field, $sum;
      $sum = 0;
    }
    print "\n";
  }

  return;
}

sub num_publishers
{ my $activ = 0;
  my $overal = @grp_rep;

  foreach my $rec (@grp_rep)
  { if ($rec->{'часы'} > 0) 
    { $activ++;}
  }
  print "\n";
  printf "общее   : %5d\n", $overal;
  printf "активных: %5d\n", $activ;
  print "\n";

  return;
}

sub print_by_assignment
{ my $rec;
  my $isfound = 0;
  my $assignment;

  print color('red');
  print "Выбери цифру: 1-возвещатель; 2-подсобный пионер; 3-пионер >_";
  print color('reset');
  my $user_inp = <STDIN>; chomp($user_inp);
  if ($user_inp == 1)
  { $assignment = 'Возвещатель';}
  if ($user_inp == 2)
  { $assignment = 'Подсобный пионер';}
  if ($user_inp == 3)
  { $assignment = 'Пионер';}

  foreach (@grp_rep)
  { if ($_->{'назначение'} eq $assignment)
    { $rec = $_;
      print_fields_of_hash($rec);
      $isfound = 'true';
    }
  }
  if (!$isfound)
  { print color('yellow');
    print "         Не найдено\n";
    print color('reset');
  }
  sums($assignment);
  return;
}

=begin
sub find_grp_number
{ #open(IN, "<$in_file") or warn "Cannot open $in_file\n";
  #my @text = <IN>;
  foreach $line (@text)
  { if ($line =~ m/group #?[0-9][0-9]*/)
    { print "The line is finded\n";};
  }
  #close(IN);
}
=end

