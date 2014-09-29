#!/usr/bin/perl -w

my $filename = $ARGV[0] || 'output.xlsx';
my $top_rows = 20;
use Excel::Writer::XLSX;
use utf8;

# Create a new Excel workbook
my $workbook = Excel::Writer::XLSX->new($filename) or die $!;
#  Add and define a format
my $header_format = $workbook->add_format(); $header_format->set_bold(); $header_format->set_align( 'center' );


# Add a worksheet headers format

my $all_vkl = $workbook->add_worksheet('Весь траффик');    
my @all_header = qw/Время ИсходящийIP Имя ВходящийIP Имя ИсходящийПорт ВходящийПорт Байт/;
$all_vkl->write_row  ('A1', \@all_header,   $header_format);
$all_vkl->set_column(0, 13, 15); # column width

my @ip_header = qw/ИсходящийIP Имя ВходящийIP Имя Mb/; 
my $ip_vkl = $workbook->add_worksheet('TOP IP');    
$ip_vkl->write_row  ('A1', \@ip_header,   $header_format); 
$ip_vkl->set_column(0, 13, 15); # column width 
my %IP;

# read ubm from stdin
my $i = 2; # insert content from this row
while (<STDIN>)
{    chomp; 	# print '.';
	 my ($bytes, $src, $dst, $port1, $port2, $time) = m/h:(\d+).*A:(\S+).*B:(\S+).*a:(\S+).*b:(\S+).*E:(..........)/;
	 next unless length "$bytes";
	 next unless length "$src"; 	 
	 next unless length "$dst"; 
	 $IP{$src.'_'.$dst} += $bytes;
	 $time =~ s/(....)(..)(..)(..)/$1-$2-$3 $4:00/;
	 my $name1 = resolve($src);
	 my $name2 = resolve($dst);
	 my @row = ($time, $src, $name1, $dst, $name2, $port1, $port2, $bytes);
	 $all_vkl->write ('A'.$i++, \@row) ;
}


# group by ip src_dst 
my $y = 2;
foreach my $src_dst (sort_by_val(\%IP, keys %IP))
{
	my ($src, $dst) = split /_/, $src_dst;
	my @row = ($src, resolve($src), $dst, resolve($dst), $IP{$src_dst}/1024/1024);
	$ip_vkl->write('A'.$y++, \@row);
	break if ($y > $top_rows + 2);
}



$workbook->close() or die "Error closing file: $!";
sub resolve
{
	my $ip = shift;
	if ((1+length $ip) < 5) {return ''}
	if ( defined $RESOLVE{$ip} ) 
		{ return $RESOLVE{$ip} } # cached
	$RESOLVE{$ip} = `host $ip|head -1| cut -f 5 -d ' ' | sed 's/\.\$//'`;
	#warn substr($RESOLVE{$ip}, 0, 3);
	if ( substr($RESOLVE{$ip}, 0, 3) eq '3(N' )
		{ $RESOLVE{$ip} = '' } # не шмогла
	return $RESOLVE{$ip};
}



sub sort_by_val
{
	my $hash = shift;
	return  sort { $hash->{$a} cmp $hash->{$b} } @_ ;
}
