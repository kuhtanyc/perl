##################################################
#
# get_ldap.pl
#
# Devleoper: D.Kuhta/2005
#
# Descritpion: returns a single user's attribute
# from the directory or return all attributes.
#
##################################################
use strict;
use gblconfig;
use Net::LDAP;

my $user=$ARGV[0] or die
  "\nusage: perl quickLDAP.pl [user_name] [attribute]\n";

my $value=$ARGV[1];

getDirectory($user,$value);

sub getLDAPLogin
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

sub LDAPConn($@)
{
  my($user,@ldapVars) = @_;

  my $ldap;
  for my $i (0..$#ldapVars)
  {
      #connect to directory server
      $ldap = Net::LDAP->new($ldapVars[$i]{server},version=>3)
      or die "unable to connect to $ldapVars[$i]{server}: $@";

      #bind user dn and password to directory
      $ldap->bind($ldapVars[$i]{dn},password=>$ldapVars[$i]{pass})
      or die "unable to bind to $ldapVars[$i]{server}: $@";
  }

  my $mesg = $ldap->search(base  => "dc=foo,dc=bar",
                           filter => "sAMAccountName=$user"
                          );

  $mesg->code() && die $mesg->error;
 
  return $mesg;
}

sub getDirectory($;$)
{
  my($user,$value) = @_;

  my @ldapVars = getLDAPLogin();

  my $mesg = LDAPConn($user,@ldapVars);

  my $max = $mesg->count;

  my @dirInfo;
  for (my $i=0; $i<$max; $i++)
  {
      my $entry = $mesg->entry($i);
 
      if ($value ne '')
      {
         push @dirInfo,{attr => $entry->get_value($value)};
        
	 print"\n";
        
	 for my $i (0..$#dirInfo) 
	 {
	    print"$value :
	    $dirInfo[$i]{attr}\n";
         }
      }
      else 
      {
         foreach my $attr ($entry->attributes)
         {
            my @values=$entry->get($attr);
            print"$attr : @values\n";
         }
      }
   }
}
