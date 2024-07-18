#######################################################
#
# ldapAdmin.pl
#
# Developer: D.Kuhta/2005
#
# Description: This utility retrieves user info from
# a directory server and displays the results in html.
# The user info. can then be modified and inserted into
# a database.
#
# Dependencies:
# -Net::LDAP - ldap perl module
# -foo.xml - contains ldap connection values
# -ldapAdmin.tmpl - html template
#
#######################################################
use strict;
use CGI qw(:standard);
use Time::localtime;
use Net::LDAP;

my $q = new CGI;

my $tmpl      = q(c:\ldapAdmin.tmpl);
my $logFile    = q(c:\ldapAdmin_log.txt);
my $insertUser = $q->param("insertUsers");
my $newUser    = $q->param("newUsersValue");

print qq{Content-type: text/html\n\n};
#-------------------------------------------------------
#The following query compares two tables to determine
#new users.
#-------------------------------------------------------
my $allUserSql = qq(SELECT distinct a.user
                    FROM foo a,
                         bar i
                    WHERE
                         upper(a.user)=upper(i.user(+))
                    AND (i.user is null));

#-------------------------------------------------------
#The $getNewUsers sql statement is passed to getNewUsers
#which executes the query and returns the results to
#@newUsers.
#-------------------------------------------------------
my $sql;
my @newUsers=getNewUsers($allUserSql);

my $log;
unless ($insertUser==1)
{
  $log='New_User_Query';
  auditLog($logFile,$log,@newUsers);
}

my $count = @newUsers;
my $aCount = $count-2;
my $mesg = "$aCount new users found.";

#-------------------------------------------------------
#@newUsers is passed, along with the server name,
#bind DN, and DN password to getDirectory, which then
#connects to the directory server to retrieve the
#the user's accountName, cn and distinguishedName.
#-------------------------------------------------------
my @userInfo=();
@userInfo = getDirectory(@newUsers);

unless ($insertUser==1)
{
  $log='User_Dir_Info';
  auditLog($logFile,$log,@userInfo);
}

#-------------------------------------------------------
#If the length of @newUsers is 0 then the following
#message is displayed in the "Results Message"
#of the html interface. showResults() opens the template
#file and places the message between the $tag value.
#-------------------------------------------------------
my $tag;
if ($aCount == 0)
{
  $tag = "showMesg";
  showResults($tmpl,$tag,"No new users found.");
  exit;
} 

#-------------------------------------------------------
#When the "Add Users to Database" button is clicked,
#@userInfo, which contains the info retrieved from the
#directory server, is combined with the html form info
#into a new AoH called @newInfo.
#-------------------------------------------------------
if ($insertUser == 1)
{
  my $num=0;
  my @newInfo;
  my $foo1;
  my $foo2;

  for my $i (0..$#userInfo)
  {
      $foo1 = $q->param("foo1_$num");
      $foo2 = $q->param("foo2_$num");

      $foo2 = "no" if $report eq "";
      push @newInfo,{uid  => $userInfo[$i]{uid},
                    cn    => $userInfo[$i]{cn},
                    oc    => $userInfo[$i]{dn},
                    foo1  => $role,
                    foo2  => $report};
      $num++;
  } 

  $log='Insert_New_Users';
  auditLog($logFile,$log,@newInfo);

  my $insert=qq(INSERT into foo
                (col1,col2,col3,col4,col5)
                values (?,?,?,?,?));

  insertUsers($insert,@newInfo);

  $tag = "<!--showMesg-->";
  showResults($tmpl,$tag,"$aCount new users successfully added to database.");
  exit;
}

$tag = "&nbsp;&nbsp;";
useTemplate($tmpl,$tag,$mesg,@userInfo);

#######################################################
=pod

=head1 sub ldapConn

=head2 Description

This function retrieves the ldap connection variables
(server name, distinguished name, password) from
foo.xml.

=head2 Parameters

n/a

=head2 Returns

@ldapVars - an AoH that contains the values needed to
connnect to the directory server.

=cut
#######################################################
sub ldapConn
{
  my $server = foo::bar("ldap","server");
  my $dn    = foo::bar("ldap","dn");
  my $pass  = foo::bar("ldap","password");

  my @ldapVars;
  push @ldapVars,{server => $server,
                  dn    => $dn,
                  pass  => $pass};

  return @ldapVars;
}

#######################################################
=pod

=head1 sub getDirectory

=head2 Description

This function connects to a directory server and
retrieves the user's accountName, cn and
distinguishedName.

=head2 Parameters

@terms - any new users

=head2 Returns

@allrows - an AoH of new user Active Directory values

=cut
#######################################################
sub getDirectory(@)
{
  my(@terms) = @_;

  my $ldap;

  my @ldapVars = ldapConn();

  for my $i (0..$#ldapVars)
  {
      #connect to directory server
      $ldap = Net::LDAP->new($ldapVars[$i]{server},version=>3)
      or die print "unable to connect to $ldapVars[$i]{server}: $@";

      #bind user dn and password to directory
      $ldap->bind($ldapVars[$i]{dn},password=>$ldapVars[$i]{pass})
      or die print "unable to bind to $ldapVars[$i]{server}: $@";
  }
  my $mesg;
  my @dirInfo;

  for my $i (0..$#terms)
  {     
      $mesg = $ldap->search(base  => "dc=foo,dc=bar",
                            filter => "accountName=$terms[$i]{uid}"
                          );

      $mesg->code() && die $mesg->error;
      my $max = $mesg->count;

      for (my $i=0; $i<$max; $i++)
      {
        my $entry = $mesg->entry($i);

        push @dirInfo,{uid  => $entry->get_value('accountName'),
                        cn  => $entry->get_value('cn'),
                        dn  => $entry->get_value('distinguishedName')
                      };
      }

      #extract office code from the dn string
      $dirInfo[$i]{dn} =~ s/OU=(.*?),OU=//gs;

      #upper case office codes
      $dirInfo[$i]{dn} = uc($1);

      #upper case full name
      $dirInfo[$i]{cn} = uc($dirInfo[$i]{cn});

      #remove non-words
      $dirInfo[$i]{cn} =~ s/\W/ /g;
  }

  $ldap->unbind();
  return @dirInfo;
}

#######################################################
=pod

=head1 sub getNewUsers

=head2 Description

This function connects to the database to find any new

users.

=head2 Parameters

$sql - the sql statement to look for new users

=head2 Returns

@newUsers - an AoH of any new users

=cut
#######################################################
sub getNewUsers($)
{
  my($sql) = @_;

  my @newUsers;
  my $sqlName="select";

  my $sth = doPrepare($sql,$sqlName);

  $sth->execute || die showError($sql,"Get New Users");
  while (my @row = $sth->fetchrow_array)
  {
      push @newUsers,{uid => $row[0]}
  }
  return @newUsers;
}

#######################################################
=pod

=head1 sub insertUsers

=head2 Description

This function connects to the database to insert the
new users and their directory info into the database.

=head2 Parameters

$sql - the insert statement
@terms - AoH of new user values

=head2 Returns

n/a

=cut
#######################################################
sub insertUsers($@)
{
  my($sql,@terms) = @_;
  my $sqlName="insert";
  my $sth = doPrepare($sql,$sqlName);

  for my $i (0..$#terms)
  {
      unless ($terms[$i]{uid} eq '')
      {     
        $sth->execute($terms[$i]{uid},
                      $terms[$i]{cn},
                      $terms[$i]{oc},
                      $terms[$i]{foo},
                      $terms[$i]{bar})
        || die showError($sql,"Insert Users");     

      }
  } 
}

#######################################################
=pod

=head1 sub useTemplate

=head2 Description

This function uses the html template to display the
user interface and query results.

=head2 Parameters

$sql - the insert statement
@terms - AoH of new user values

=head2 Returns

n/a

=cut
#######################################################
sub useTemplate($$$@)
{
  my($tmpl,$tag,$mesg,@data) = @_;

  open(FILE, $tmpl) or die print "Can't open file: $!\n";

  my @html = <FILE>;

  foreach my $line (@html)
  {
      $line =~ /<!--showMesg-->/ ? print "$mesg\n" : undef;

      if($line=~/$tag/)
      {
        my $num=0;
        for my $i (0..$#data)
        {
            my $color="#FFFFFF";
            my $select_user;
            my $select_admin;
            my $checked = "checked";
            $data[$i]{dn} eq 'foo' ? $color="yellow" : $color;             
            $data[$i]{dn} eq 'foo' ? $select_admin="selected" : $select_user="selected";
            $data[$i]{dn} eq 'foo' ? $checked="" : $checked;

            unless ($data[$i]{uid} eq '')
            {
              $line .= qq(
              <tr bgcolor=$color>
                <td><center>$data[$i]{uid} </td>
                <td><center>$data[$i]{cn} </td>
                <td><center>$data[$i]{dn} </td>
                <td>
                <center>
                <select name="role_$num">
                  <option value="User" $select_user>User</option>
                  <option value="Admin" $select_admin>Admin</option>
                  <option value="Support">Support</option>
                  <option value="Training">Training</option>
                </select>
                </td>
                <td>
                <center>
                <input name="report_$num" type="checkbox" value="yes" $checked>
                </td>
              </tr>\n);
            }
            $num++;
        }
      }
      print "$line";
  }
}

#######################################################
=pod

=head1 sub showResults

=head2 Description

This function uses the html template to display any
results comments.

=head2 Parameters

$tmpl - the template file
$tag - location in the template to display message
$mesg - the message to display

=head2 Returns

n/a

=cut
#######################################################
sub showResults($$;$)
{
  my($tmpl,$tag,$mesg) = @_;

  open(FILE, $tmpl)
  or print "Can't open file: $!\n";
  my @html = <FILE>;

  my $last = (<FILE>)[-1];

  foreach my $line (@html)
  {
      $line =~ /<!--showMesg-->/ ? $line .= "$mesg\n" : undef;
      $line =~ /<!--showTable-->/ ? $line .= "<!--\n" : undef;
      $line =~ /<\/html>/ ? $line .= "-->\n" : undef;
      print "$line";
  }
}

#######################################################
=pod

=head1 sub doPrepare

=head2 Description

This function prepares a sql statment and return a
handle.

=head2 Parameters

$sql - the sql statement to prepare
$sqlName - required value

=head2 Returns

n/a

=cut
#######################################################
sub doPrepare($$)
{
  my($sql,$sqlName) = @_;
  my $sth = foo::bar($sqlName,$sql)
      or die print "Prepare failed, $DBI::errstr\n$sql\n"; 
  return $sth;
}

#######################################################
=pod

=head1 sub showError

=head2 Description

This function displays an error message for the sql
queries. It shows which function caused the error and
the actual sql statement.

=head2 Parameters

$sql - the sql statement to prepare
$desc - displays the fuction where the sql failed

=head2 Returns

n/a

=cut
#######################################################
sub showError($$)
{
  my($sql,$desc) = @_;
  print qq(<b>Program Error - [Function: $desc]</b><br>Could not execute SQL statement: $sql);
}

#######################################################
=pod

=head1 sub auditLog

=head2 Description

This function logs query and insert activity and info.

=head2 Parameters

$logFile - the log file
$log - determines which activity to log
@data - user data to log

=head2 Returns

n/a

=cut
#######################################################
sub auditLog($$@)
{
  my($logFile,$log,@data) = @_;
  open(FILE, ">>$logFile")
  or print "Can't open file: $!\n";

  my $tm = localtime;
  my $m  = $tm -> mon+1;
  my $d  = $tm -> mday;
  my $y  = $tm -> year+1900;
  my $hr  = $tm -> hour;
  my $min = $tm -> min;
  my $sec = $tm -> sec;

  print FILE "$log [$m/$d/$y - $hr:$min:$sec]\n";
  print FILE "---------------------------------------\n";

  if ($log eq 'New_User_Query')
  {
      for my $i (0..$#data)
      {
        print FILE "$data[$i]{uid}\n";
      }
  }
  elsif ($log eq 'User_Dir_Info')
  {
      for my $i (0..$#data)
      {
        print FILE "$data[$i]{uid}|$data[$i]{cn}|$data[$i]{dn}\n";
      }
  } 
  elsif ($log eq 'Insert_New_Users')
  {
      for my $i (0..$#data)
      {
        print FILE "$data[$i]{uid}|$data[$i]{cn}|$data[$i]{oc}|$data[$i]{role}|$data[$i]{rep}\n";
      }
  }
  print FILE "\n";
  close(FILE);
}
