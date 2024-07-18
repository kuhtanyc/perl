############################################
#
# webLog.pl
#
# Developer: D.Kuhta/2005
#
# Description: View IIS server logs in html
# format with the ability to view usage
# statistics for a specific application.
#
# Dependencies:
# webLog.tmpl - html template
#
############################################
use strict;
use CGI qw(:standard);

my $q = new CGI;
my $app = $q->param("app");
my $log = $q->param("log");

my $tmpl = q(c:\webLog.tmpl);

my @logData=();
if ($log)
{
  open(FILE, $log) or die print "Can't open file:
$!\n";
  @logData = <FILE>;
  close(FILE);
}

print qq{Content-type: text/html\n\n};

my $hits  = 0;
my $query = 0;
my $excel = 0;
my $batch = 0;
my $error = 0;
my $tag;
my $date;
my @data=();
my @nData=();
my @users=();

#extract lines of interest from the server log
foreach my $line (@logData)
{
  $date = $line if $line =~/#Date:/;

  if ($line != /#(.*?):/)
  {
      if ($line =~ /$app/)
      {
        #split each column by whitespace
        @data=split(/ /,$line);

        #load array data into an AoH
        push @nData,{date  => $data[0],
                      time  => $data[1],
                      user  => $data[3],
                      query => $data[8]};

        #get counts from query string
        $data[8] !~ /foo/ ? $query++ : $query;
        $data[8] =~ /bar/ ? $batch++ : $batch;

        #load users into a new array
        push(@users,$data[3]);

        $hits++;
      }
  }
}

#get total hits for each user
my %active=activeUser(@users);

#sort each key value numerically then put them
#into a new array and take the first element as
#the most active user
my @activeUser=();
foreach my $u (sort {$active{$b} <=> $active{$a}} keys
%active)
{
  my $user="$u ($active{$u})";
  push(@activeUser,$user);
}
my $activ=$activeUser[0];

#get number of unique users
my $users_ref=\@users;
my @uniqUsers=uniqUsers(@$users_ref);
my $users=@uniqUsers;

useTemplate($date,$hits,$query,$batch,$excel,$error,$users,$activ,$app,@nData);

sub useTemplate
{
 
my($date,$hits,$query,$batch,$excel,$error,$users,$activ,$app,@nData)
= @_;

  my @html = getTemplate($tmpl);

  foreach my $line (@html)
  {
      if    ($line=~/<!--date-->/)  { $line .=
qq($date);  }
      elsif ($line=~/<!--hits-->/)  { $line .=
qq($hits);  }
      elsif ($line=~/<!--query-->/) { $line .=
qq($query); }
      elsif ($line=~/<!--batch-->/) { $line .=
qq($batch); }
      elsif ($line=~/<!--excel-->/) { $line .=
qq($excel); }
      elsif ($line=~/<!--error-->/) { $line .=
qq($error); }
      elsif ($line=~/<!--users-->/) { $line .=
qq($users); }
      elsif ($line=~/<!--activ-->/) { $line .=
qq($activ)||'n/a'; }
      elsif ($line=~/<!--app-->/)  { $line .=
qq($app);  }
      elsif ($line=~/<!--logData-->/)
      {
        for my $i (0..$#nData)
        {
            $line .= qq(<tr>
                        <td><font
size=1>$nData[$i]{date}</td>
                        <td><font
size=1>$nData[$i]{time}</td>
                        <td><font
size=2>$nData[$i]{user}</td>
                        <td><font
size=2>$nData[$i]{query}</td>
                        </tr>);
        }
      } 
      print qq($line);
  }
}

sub activeUser
{
  my(@users) = @_;
  my %active=();
  foreach my $item (@users)
  {
      $active{$item}++;
  }
  return %active;
}

sub uniqUsers
{
  my(@users) = @_;
  my %seen=();
  my @uniq=(); 
  foreach my $item (@users)
  {
      push(@uniq,$item) unless $seen{$item}++;
  }
  return @uniq;
}

sub getTemplate
{
  open(FILE, $tmpl) or die print "Can't open file:
$!\n";
  my @html = <FILE>;
  close(FILE);
  return @html;
}
