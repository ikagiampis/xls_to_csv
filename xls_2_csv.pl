#For bugs: ikagiampis@gmail.com Enjoy!
use strict;
use warnings;
use Data::Dumper;
use Getopt::Std;
use Spreadsheet::ParseExcel;

$| = 1;

sub main {
	my %opts;

	#d: directory

	getopts( 'd:', \%opts );

	if ( !checkusage( \%opts ) ) {
		usage();
		exit();
	}

	my $input_dir = $opts{"d"};

	my $gene_number = $opts{"l"};

	my @files = get_files($input_dir);

foreach my $file (@files){
	my $file_path = "$input_dir".'/'."$file";
	xls_to_csv($file_path)
}
}

#----------------------------------------------------------
sub xls_to_csv{
	
my $sourcename = shift or die "invocation: $0 <source file>\n";
my $source_excel = Spreadsheet::ParseExcel->new();
my $source_book = $source_excel->Parse("$sourcename") or die "Could not open source Excel file $sourcename: $!";;
my $storage_book;


foreach my $source_sheet_number (0 .. $source_book->{SheetCount}-1)
{
 my $source_sheet = $source_book->{Worksheet}[$source_sheet_number];

 print "--------- SHEET:", $source_sheet->{Name}, "\n";

 next unless defined $source_sheet->{MaxRow};
 next unless $source_sheet->{MinRow} <= $source_sheet->{MaxRow};
 next unless defined $source_sheet->{MaxCol};
 next unless $source_sheet->{MinCol} <= $source_sheet->{MaxCol};

 foreach my $row_index ($source_sheet->{MinRow} .. $source_sheet->{MaxRow})
 {
  foreach my $col_index ($source_sheet->{MinCol} .. $source_sheet->{MaxCol})
  {
   my $source_cell = $source_sheet->{Cells}[$row_index][$col_index];
   if ($source_cell)
   {
    #print "( $row_index , $col_index ) =>", $source_cell->Value, "\t";
    print  $source_cell->Value, ",";
   }
  } 
  print "\n";
 } 
}
print "done!\n";
}
#----------------------------------------------

sub get_files {

	my $input_dir = shift;

	unless ( opendir( INPUTDIR, $input_dir ) ) {
		die "\n Unable to open durectory '$input_dir'\n";
	}

	my @files = readdir(INPUTDIR);

	#print Dumper( \@files );

	closedir(INPUTDIR);

	@files = grep( /\.xlsx*$/i, @files );

	return @files;
}

#----------------------------------------------------------
sub checkusage {
	my $opts = shift;

	#d: directory

	my $directory = $opts->{"d"};

	unless ( defined($directory) ) {
		return 0;
	}

	return 1;
}

#----------------------------------------------
sub usage {

	print <<USAGE;
	
    usage: perl xls_to_csv.pl <options>
	-d <directory>
	
USAGE
}
#----------------------------------------------
main();
