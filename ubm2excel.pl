#!/usr/bin/perl -w

my $filename = $ARGV[0] || 'ubm2excel.xlsx';

use Excel::Writer::XLSX;

# Create a new Excel workbook
my $workbook = Excel::Writer::XLSX->new($filename) or die $!;
#  Add and define a format
my $header_format = $workbook->add_format(); $header_format->set_bold(); $header_format->set_align( 'center' );

# Add a worksheet
my $time_vkl = $workbook->add_worksheet('Время');        $time_vkl->write_row(1, qw/ Время Мбайт /, $header_format ); 
my $ip_vkl   = $workbook->add_worksheet('IP адреса');    $ip_vkl->write_row(1, qw/ ИсходящийIP Имя ВходящийIP Имя Мбайт /, $header_format);
my $port_vkl = $workbook->add_worksheet('Порты');        $port_vkl->write_row(1, qw/Время Мбайт/, $header_format); 

my (%TIME, %IP, %PORT);

# read ubm from stdin
while (<>)
{    chomp;
	 my ($bytes, $src, $dst, $port1, $port2, $time) = m/h:(\d+).*A:(\S+).*B:(\S+).*a:(\S+).*b:(\S+).*E:(..........)/;

	 $IP{$src."\t".$dst} += $bytes;
	 $PORT{$port1} += $bytes;
	 $TIME{$time} += $bytes;
}

my $ip_arr = hashref2arrayref(\%IP);
my $port_arr = hashref2arrayref(\%PORT);
my $time_arr = hashref2arrayref(\%TIME);



sub hashref2arrayref
{
	my $hash = shift;
	my $arrayref = [];
	foreach my $key (keys %$hash)
	{
		push [ $key, $hash->{$key}], @$arrayref;
	}
	return $arrayref;
}

sub sort_print_hash
{
	my ($hash, $worksheet) = @_;
	my ($row, $col) = (3, 1);
	foreach my $key ( sort hash_comp (keys %$hash))
	{
		$worksheet->write(1, $key);
        print "$key\t$TRAF{$key}\n";
	}
}

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




