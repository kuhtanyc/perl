#######################################################
#
# writeExcel.pl
#
# Developer: D.Kuhta/2005
#
# Description: This program performs keyword filtering
# on a set of data and then saves and formats the
# results to an Excel spreadsheet.
#
#######################################################

use strict;
use Spreadsheet::WriteExcel;

my $dir = qq(C:\\Inetpub\\Excel);
my $file = "foo";

my @data=();
foreach my $line (<DATA>)
{
  my @nData=split(/\t/,$line);

  push @data,{name => $nData[0],
              race => $nData[1],
              age  => $nData[2],
              misc => $nData[3]};
}

writeExcel($file,$dir,@data);

sub writeExcel($$@)
{
  my($file,$dir,@data) = @_;

  #create a new workbook and worksheet
  my $workbook  =
Spreadsheet::WriteExcel->new("$dir/$file".".xls")
  or print "Can't create workboook: $!\n";

  my $worksheet = $workbook->add_worksheet();

  #set column name format
  my $col_format = $workbook->add_format();
  $col_format->set_bold();

  #set String format
  my $String_format = $workbook->add_format();
  $String_format->set_align('left');

  #set Age format
  my $Age_format = $workbook->add_format();
  $Age_format->set_num_format('000');
  $Age_format->set_align('left');

  # write column names
  $worksheet->set_column(0,0,20);
  $worksheet->write(0,0,"Count",$col_format);
  $worksheet->set_column(0,1,20);
  $worksheet->write(0,1,"Name",$col_format);
  $worksheet->set_column(0,2,20);
  $worksheet->write(0,2,"Race",$col_format);
  $worksheet->set_column(0,3,20);
  $worksheet->write(0,3,"Age",$col_format);
  $worksheet->set_column(0,4,20);
  $worksheet->write(0,4,"Misc",$col_format);

  my $row = 1;
  my $count = 1;

  foreach my $i (0..$#data)
  {
      foreach my $role (keys %{$data[$i]})
      {
        $data[$i]{$role} =~
s/(^|;)\s*(na|n\/a|unknown.*?)\s*(;|$)/$1$3/ig;
      }

      #set age to zero
      if ($data[$i]{age} eq "") {$data[$i]{age} = 0;}

      #eliminate non-numeric characters
      $data[$i]{age}  =~ s/\D//g;

      #concatenate a blank space so excel reads this
value as text,
      #otherwise preceding and trailing zeros are lost
      $data[$i]{misc} = $data[$i]{misc}." ";

      #remove the "0" so Excel doesn't make age
"00000000"
      if ($data[$i]{age} == 0) {$data[$i]{age} =~
tr/0//d;}

      #apply keyword filter to data
      doFilter(\$data[$i]{race});

      #eliminate repeating numbers
      $data[$i]{age}=""if(substr($data[$i]{age},0,5)
eq
substr($data[$i]{age},0,1)x5);
      $data[$i]{misc}=""if(substr($data[$i]{misc},0,5)
eq
substr($data[$i]{misc},0,1)x5);

      #apply elimination requirements on misc numbers
      $data[$i]{misc} = miscNum($data[$i]{misc});

      #write all filtered data to appropriate cells
      $worksheet->write($row,0,$count,       
$String_format);
   
$worksheet->write($row,1,$data[$i]{name},$String_format);
   
$worksheet->write($row,2,$data[$i]{race},$String_format);
      $worksheet->write($row,3,$data[$i]{age},
$Age_format);
   
$worksheet->write($row,4,$data[$i]{misc},$String_format);

      $row++;
      $count++;
  } 
}

sub doFilter($$)
{
  my ($race_ref) = @_;

  foreach my $term (getTerms())
  {
      while ($$race_ref =~ /\b$term\b/ig){$$race_ref =
'';}
   
      $$race_ref =~ s/\W/ /g;          #replace
nonwords with a space
      $$race_ref =~ s/\S*\d\S*/ /g;    #replace words
with nums with a
space
      $$race_ref =~ s/\b$term\b/ /ig;  #replace filter
terms with a
space
      $$race_ref =~ s/\s+/ /g;        #collapse
whitespace to one
space
      $$race_ref =~ s/(^\s+)|(\s$)//g; #trim
leading/trailing
whitespace
  }
}

sub miscNum
{
  my ($miscNum) = @_;

  #no less than 5 characters
  if(length($miscNum) < 5) {
      $miscNum = "";
  }
  else {
      #remains alphanumeric
      if($miscNum =~ /[0-9]/ && $miscNum !~
/[A-Za-z]/) {
        #if a numeric value, must not be greater
        #than 9999, ie: 00000009999
        unless($miscNum > 9999) {
            $miscNum = "";
        }
      }
  }
  return $miscNum;
}

sub getTerms {
  return qw(dwarf human orc);
}

__DATA__
bilbo baggins    hobbit    144    AA45000
frodo baggins    hobbit    111    00000A996
sam gamgee    hobbit    n/a    TX003433
gimli    dwarf    200    4873
gandalf    human    unknown    00000009999
