################################################
#
# date_diff.pl
#
# Description: this program gets the difference
# of two dates.
#
################################################
use Time::Local;

$startdate = '07012005';
$enddate  = '07042005';

$startsec = dateToSec($startdate);
$endsec = dateToSec($enddate);
$diff = ($startsec-$endsec)/(1000*3600*24)*-1000;

print "$diff\n";

if($diff>60){print "error: interval is greater than 60
days\n";}
else{print "60 or less\n";}

sub dateToSec
{
  $date = shift;
  $_ = $date;
  my($m, $d, $y) = /(\d{2})(\d{2})(\d{4})/;
  $m--;                                    #months
are 0-based
  my $time = timegm(0, 0, 0, $d, $m, $y);  #GMT is
DST-independent
  #print scalar localtime($time);
  return $time;
}
