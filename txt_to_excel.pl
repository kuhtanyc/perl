use Win32::OLE;
use strict;

#collect data from txt file
open(FILE,"testdata.txt")||die "can't open file:
$!\n";

excelFormat();

###################################################
=item excelFormat

This function ingests data in a delimited format,
applies specified string elimination regexes to the
data and then saves the formatted results to an
Excel spreadsheet using OLE.

=cut
###################################################
sub excelFormat() {

  #output directory of the Excel spreadsheet
  my $outputdir ='C:\Documents and Settings\kuhtad\Desktop\OLE\output.xls';

  #use existing instance if Excel is already running
  my $ex;
  eval {$ex = Win32::OLE->GetActiveObject('Excel.Application')};
  die "Excel not installed" if $@;
       
  unless (defined $ex) {
      $ex = Win32::OLE->new('Excel.Application', sub{$_[0]->Quit;})
      or die "cannot start Excel";
  }

  #open a new workbook
  my $book;
  $book = $ex->Workbooks->Add;

  #write to a particular cell
  my $sheet;
  $sheet = $book->Worksheets(1);

  #create the required column names
  $sheet->Cells(1,1)->{Value} = 'Name'; 
  $sheet->Cells(1,2)->{Value} = 'SSN';
  $sheet->Cells(1,3)->{Value} = 'DOB';
  $sheet->Cells(1,4)->{Value} = 'Phone';
  $sheet->Cells(1,5)->{Value} = 'Misc_Num';

  #initialize the rows
  my $row=2;

  #write each record into the worksheet
  my $line;
  while($line = <FILE>) 
  {
      chomp; 
      #separate each cell value
      my $name;my $ssn;my $dob;
      my $phone;my $misc;
   
      ($name,$ssn,$dob,$phone,$misc)=split("\t",$line);
 
      #write each row to the worksheet
      foreach($line)
      {
         #perform regexes on $name to eliminate specified strings;
         #eliminate instances of "n/a","na","unknown","not known",
         $name =~ s/na//g;
         $name =~ s/n\/a//g;
         $name =~ s/unknown//g;
         $name =~ s/not known//g;
         #write filtered data to the appropriate cell
         $sheet->Cells($row,1)->{Value} = $name;

         #eliminate any alphabetic strings and repeating
         #numbers (ie: 11111,22222,33333,etc...)
         $ssn =~ s/[A-Za-z]/ /g;
         $ssn =~ s/[0-9]{10,}/ /g;

         #write filtered data to the appropriate cell
         $sheet->Cells($row,2)->{Value} = $ssn;

         #perform regexes on $dob to eliminate specified strings;
         #eliminate instances of "n/a","na","unknown","notknown",
         #repeating numbers (ie: 11111,22222,33333,etc...)
         $dob =~ s/na//g;
         $dob =~ s/n\/a//g;
         $dob =~ s/unknown//g;
         $dob =~ s/not known//g;
         #write filtered data to the appropriate cell
         $sheet->Cells($row,3)->{Value} = $dob;

         #eliminate any alphabetic strings and repeating
         #numbers (ie: 11111,22222,33333,etc...)
         $phone =~ s/[A-Za-z]/ /g;
         #write filtered data to the appropriate cell
         $sheet->Cells($row,4)->{Value} = $phone;

         #perform regexes on $misc to eliminate specified strings;
         #eliminate instances of "n/a","na","unknown","not known",
         #repeating numbers (ie: 11111,22222,33333,etc...)
         $misc =~ s/na//g;
         $misc =~ s/n\/a//g;
         $misc =~ s/unknown//g;
         $misc =~ s/not known//g;
         #write filtered data to the appropriate cell
         $sheet->Cells($row,5)->{Value} = $misc;

         ++$row;
      } 
      print $line;
  }
  close(FILE);
