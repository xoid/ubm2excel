#!/usr/bin/perl -w

my $filename = $ARGV[0] || 'ubm2excel.xlsx';

use Excel::Writer::XLSX;
use utf8;

# Create a new Excel workbook
my $workbook = Excel::Writer::XLSX->new($filename) or die $!;
#  Add and define a format
my $header_format = $workbook->add_format(); $header_format->set_bold(); $header_format->set_align( 'center' );

# headers
my @ip_header   = qw/ИсходящийIP Имя ВходящийIP Имя Мбайт/;
my @time_header = qw/Время Мбайт/;
my @port_header = qw/Порт  Мбайт/;
# Add a worksheet
my $ip_vkl   = $workbook->add_worksheet('IP адреса');    $ip_vkl->write_row  ('A1', \@ip_header,   $header_format);
my $time_vkl = $workbook->add_worksheet('Время');        $time_vkl->write_row('A1', \@time_header, $header_format ); 
my $port_vkl = $workbook->add_worksheet('Порты');        $port_vkl->write_row('A1', \@port_header, $header_format); 
$ip_vkl->set_column(0, 3, 40); # first 4 column width set to 40

my (%TIME, %IP, %PORT);

# read ubm from stdin
while (<>)
{    chomp;
	 my ($bytes, $src, $dst, $port1, $port2, $time) = m/h:(\d+).*A:(\S+).*B:(\S+).*a:(\S+).*b:(\S+).*E:(..........)/;
	 $time =~ s/(....)(..)(..)(..)/$1-$2-$3 $4:00/;
	 $IP{$src."_".$dst} += $bytes;
	 $PORT{$port1} += $bytes;
	 $TIME{$time} += $bytes;
}


my $ip_arr   = hashref2arrayref(\%IP);
my $port_arr = hashref2arrayref(\%PORT);
my $time_arr = hashref2arrayref(\%TIME);

sub hashref2arrayref
{
	my $hash = shift;
	my $arrayref = [];
	foreach my $key (keys %$hash)
	{
		#warn '.';
		$hash->{$key} /= 1024*1024;
		my @keys = split /_/, $key;
		push @$arrayref, [ @keys, $hash->{$key} ];
	}
	# sort by last value
	sort { $a->[-1] <=> $b->[-1] } @$arrayref;
	return $arrayref;
}

# write to worksheets
$ip_vkl->write  ('B1', hashref2arrayref(\%IP));
$time_vkl->write('B1', hashref2arrayref(\%TIME));
$port_vkl->write('B1', hashref2arrayref(\%PORT));

#    $format->set_color( 'red' );
#    $format->set_align( 'center' );

# Write a formatted and unformatted string, row and column notation.
#$col = $row = 0;
#$worksheet->write( $row, $col, 'Hi Excel!', $format );
#$worksheet->write( 1, $col, 'Hi Excel!' );

# Write a number and a formula using A1 notation
#$worksheet->write( 'A3', 1.2345 );
#$worksheet->write( 'A4', '=SIN(PI()/4)' );

#sub array2excel
#{
#	
#}

sub hash_comp {   $TRAF->{$b} <=> $TRAF->{$a}   }

#sub cp1251 { encode('cp1251', shift) }

my %RESOLVE;

sub resolve_ip
{
	my $ip = shift;
	if ($RESOLVE{$ip} ne undef) { return $RESOLVE{$ip} }
	$RESOLVE{$ip} = `host $ip|head -1| cut -f 5 -d ' ' | sed 's/\.$//'`;
	if ($RESOLVE{$ip} eq '' or $RESOLVE{$ip} eq undef or $RESOLVE{$ip} eq '3(NXDOMAIN)')
		{ return 'не определить' }
	return $RESOLVE{$ip};
}


