#!/usr/bin/perl -w

use strict;
my $filename = $ARGV[0] || 'output.xlsx';
my $top_rows = 20;

use Socket;
use Excel::Writer::XLSX;
use utf8;

my %RESOLVE;
# Create a new Excel workbook
my $workbook = Excel::Writer::XLSX->new($filename) or die $!;
#  Add and define a format
my $format = $workbook->add_format(); $format->set_num_format('0.000');
my $header_format = $workbook->add_format(); $header_format->set_bold(); $header_format->set_align( 'center' );

my $all_vkl = $workbook->add_worksheet('Весь траффик');    
my @all_header = qw/Время ИсходящийIP Имя ВходящийIP Имя ИсходящийПорт ВходящийПорт Байт/;
$all_vkl->write_row  ('A1', \@all_header,   $header_format);
$all_vkl->set_column(0, 13, 15); # column width

my @ip_header = qw/ИсходящийIP Имя ВходящийIP Имя Mb/; 
my $ip_vkl = $workbook->add_worksheet('TOP IP');    
$ip_vkl->write_row  ('A1', \@ip_header,   $header_format); 
$ip_vkl->set_column(0, 13, 15); # column width 
my %IP;

my @day_header = qw/Дата Mb/; 
my $day_vkl = $workbook->add_worksheet('TOP date');    
$day_vkl->write_row  ('A1', \@day_header,   $header_format); 
$day_vkl->set_column(0, 13, 15); # column width 
my %DAY;

# read ubm from stdin
my $i = 2; # insert content from this row
while (<STDIN>)
{    chomp; 	# print '.';
	 my ($bytes, $src, $dst, $port1, $port2, $time) = m/h:(\d+).*A:(\S+).*B:(\S+).*a:(\S+).*b:(\S+).*E:(..........)/;
	 next unless length "$bytes";
	 next unless length "$src"; 	 
	 next unless length "$dst"; 
	 $time =~ s/(....)(..)(..)(..)/$1-$2-$3 $4:00/;
	 my $name1 = resolve($src);
	 my $name2 = resolve($dst);
	 my @row = ($time, $src, $name1, $dst, $name2, $port1, $port2, $bytes);
	 $all_vkl->write ('A'.$i++, \@row) ;
	 $IP{$src.'_'.$dst} += $bytes;
	 #warn "";
	 $DAY{"$1-$2-$3"} += $bytes;
}

# group by ip src_dst 
my $y = 2;
foreach my $src_dst (sort_by_val(\%IP))
{
	my ($src, $dst) = split /_/, $src_dst;
	#warn $src_dst, $IP{$src_dst};
	my @row = ($src, resolve($src), $dst, resolve($dst), $IP{$src_dst}/1024/1024);
	$ip_vkl->write('A'.$y++, \@row, $format);
	last if ($y > $top_rows + 2);
}

$y = 2;
foreach my $day (sort_by_val(\%DAY))
{
	my @row = ($day, $DAY{$day}/1024/1024);
	$day_vkl->write('A'.$y++, \@row);
	last if ($y > $top_rows + 2);
}


$workbook->close() or die "Error closing file: $!";
sub resolve
{	#return 'lol';
	my $ip = shift;
	if ((1+length $ip) < 5) {return ''}
	if ( defined $RESOLVE{$ip} ) 
		{ return $RESOLVE{$ip} } # cached
	$RESOLVE{$ip} = gethostbyaddr( inet_aton($ip), AF_INET) || ' ';
	#warn $RESOLVE{$ip};
	#warn substr($RESOLVE{$ip}, 0, 3);
	if ( substr($RESOLVE{$ip}, 0, 3) eq '3(N' )
		{ $RESOLVE{$ip} = '' } # не шмогла
	return $RESOLVE{$ip};
}



sub sort_by_val
{
	my $hash = shift;
	use Data::Dumper;
	#warn Dumper 'hash ', $hash, 'keys ';
	return  sort { $hash->{$b} <=> $hash->{$a} } keys %$hash;
}
