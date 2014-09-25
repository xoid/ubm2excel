#!/usr/bin/perl -w

# comment 1
# comment 2
# comment 3

my $filename = $ARGV[0] || 'output.xlsx';
use encoding "ru_RU.CP1251";
#use Encode;

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
<<<<<<< HEAD
my $ip_vkl   = $workbook->add_worksheet('IP адреса');    $ip_vkl->write_row  ('A1', [ 'ИсходящийIP', 'Имя', 'ВходящийIP', 'Имя', 'Мбайт' ], $header_format);
my $time_vkl = $workbook->add_worksheet('Время');        $time_vkl->write_row('A1', [ 'Время', 'Мбайт' ], $header_format ); 
my $port_vkl = $workbook->add_worksheet('Порты');        $port_vkl->write_row('A1', [ 'Время', 'Мбайт' ], $header_format); 
=======
my $ip_vkl   = $workbook->add_worksheet('IP адреса');    $ip_vkl->write_row  ('A1', \@ip_header,   $header_format);
my $time_vkl = $workbook->add_worksheet('Время');        $time_vkl->write_row('A1', \@time_header, $header_format ); 
my $port_vkl = $workbook->add_worksheet('Порты');        $port_vkl->write_row('A1', \@port_header, $header_format); 
$ip_vkl->set_column(0, 3, 40); # first 4 column width set to 40
>>>>>>> 5d66730b1ff8c1663d57a00cac5722d1d8c00141

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


sub hashref2arrayref
{
	my $hash = shift;
	my $arrayref = [];
	foreach my $key (keys %$hash)
	{
		#warn '.';
		$hash->{$key} /= 1024*1024;
		if (my ($src_ip, $dst_ip) = split /_/, $key) # this is ip
		{
			push @$arrayref, [ $src_ip, resolve($src_ip), $dst_ip, resolve($dst_ip), $hash->{$key} ];
		};
		push @$arrayref, [ $key, $hash->{$key} ];
	}
	# sort by last value
	return sort { $a->[-1] <=> $b->[-1] } @$arrayref;
}

# write to worksheets
$ip_vkl->write  ('A2', hashref2arrayref(\%IP));
$time_vkl->write('A2', hashref2arrayref(\%TIME));
$port_vkl->write('A2', hashref2arrayref(\%PORT));

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

<<<<<<< HEAD
sub cp1251 { encode('1251', shift) }
=======
sub resolve
{
	my $ip = shift;
	if ((1+length $ip) < 5) {return ''}
	if ( length $RESOLVE{$ip} ) { return $RESOLVE{$ip} } # cached
	$RESOLVE{$ip} = `host $ip|head -1| cut -f 5 -d ' ' | sed 's/\.\$//'`;
	if (! length $RESOLVE{$ip} or $RESOLVE{$ip} eq '3(NXDOMAIN)')
		{ return 'неизвестно' } # не шмогла
	return $RESOLVE{$ip};
}
>>>>>>> 5d66730b1ff8c1663d57a00cac5722d1d8c00141


