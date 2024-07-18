###########################################
# renames part of a filename and saves as
# new file
###########################################
use strict;
use warnings;
use File::Copy;
 
#existing file to be renamed
my $file1 = 'foo_001.txt';
 
(my $file2=$file1) =~ s/^(.*?)_/bar_/g;
 
#copy the original filename to the replaced filename
copy($file1,$file2) or die "copy failed: $!";
