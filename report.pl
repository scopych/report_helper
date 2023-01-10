#!/usr/bin/perl -T
#use warnings;
use strict;
#use utf8;
#use feature 'unicode_strings'; 

#my $in_file = "$ARGV[0]";

#open(IN, "<$in_file") or warn "Cannot open $in_file\n";

#my @text = <IN>;
my @grp_rep;
#my $grp_number;

fill_grp_rep();
#find_grp_number();

my $c = -1;
while ($c != 0)         # main loop
{ print_menu();
  print ">>>_";
  $c = <STDIN>; chomp($c);
  if ($c == 1)
  { print_grp_rep();}
  if ($c == 2)
  { print "Введи имя или часть имени_";
    print_by_name();
  }
  if ($c == 3)
  { print_by_assignment();}
  if ($c == 4)
  { sums();}
  if ($c == 5)
  { num_publishers();}
  if ($c == 0)
  { exit;}
}

#close(IN);

sub print_menu
{  print "      МЕНЮ\n";
  print "0 Выйти\n";
  print "1 Показать весь отчет\n";
  print "2 Показать отчет по имени\n";
  print "3 Показать отчет по назначению\n";
  print "5 Display sums of all fields\n";
  print "5 Number of publishers\n";
}

sub fill_grp_rep
{ my @fields = qw(имя публикации видео часы повторы изучения примечания назначение активность);#qw(name pubs video hours returns studis notes serveAs activity);
  while (<>) #foreach (@text) #while (<IN>)
  { chomp;
    my %rec;
    my @row = split(/,/, $_, 9);#split /,/;
    next if (!defined $row[0]);
    #if ($row[0] =~ m/^[А-Я][а-я]+?\s+?[А-Я][а-я]*?/)
    @rec{@fields} = @row;#{ @rec{@fields} = @row;
    push @grp_rep, \%rec;#  push @grp_rep, \%rec;
    #}
  }
  #close(IN);
}

sub print_grp_rep
{ print "Отчет группы № \n";
  foreach (@grp_rep)
  { my $rec = $_;
    print_fields_of_hash($rec);
  }
  sums();
}

sub print_fields_of_hash
{ my $rec = $_;
  print "\n";
  print "      Имя:         $rec->{'имя'}\n"; 
  print "      Публикации:  $rec->{'публикации'}\n";
  print "      Видео:       $rec->{'видео'}\n";
  print "      Часы:        $rec->{'часы'}\n";
  print "      Повторы:     $rec->{'повторы'}\n";
  print "      Изучения:    $rec->{'изучения'}\n";
  print "      Примечания:  $rec->{'примечания'}\n";
  print "      Назначение:  $rec->{'назначение'}\n";
  print "      Активность:  $rec->{'активность'}\n";
  print "\n";
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
  { print "         Не найдено\n";}
}

sub sums
{ my $rec;
  my $field;
  my @fields = qw(публикации видео часы повторы изучения); #qw(pubs video hours returns studis);
  my $sum = 0;

  print "          Общие цифры\n";
  foreach $field (@fields)
  { foreach $rec (@grp_rep)
    { $sum += $rec->{$field};}
    printf "%15s: %5d\n", $field, $sum;
    $sum = 0;
  }
  print "\n";
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
}

sub print_by_assignment
{ my $rec;
  my $isfound = 0;
  my $assignment;

  print "Выбери цифру: 1-возвещатель; 2-подсобный пионер; 3-пионер >_";
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
  { print "         Не найдено\n";}
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

