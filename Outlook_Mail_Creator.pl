#perl2exe_info FileDescription =Outlook_Mail_Creator
#perl2exe_info FileVersion=1.0
#perl2exe_info InternalName=Outlook_Mail_Creator.pl
#perl2exe_info LegalCopyright=Akulich Dmitry akulich.d@gmail.com
#perl2exe_info ProductName=Outlook Mail Creator
#perl2exe_info ProductVersion=1
use Getopt::Long;
use Cwd qw(cwd);
use strict;
use Win32::GUI();
my $ini_file="OMC.ini";
my $emailfileslist="email_templates.lst";
my %emailtmpl;
my $log="omc.log";
my $filehandle;
my $storedb_f="store.db";
my %storedb;
my %inidata;
my $tmp;
my $label;
my $itemcount=0;
my $verbose=0;
my @months = qw( Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec );
my @days = qw(Sun Mon Tue Wed Thu Fri Sat Sun);
GetOptions ('verbose' =>  \$verbose);

(my $sec,my $min,my $hour,my $mday,my $mon,my $year,my $wday,my $yday,my $isdst) = localtime();
$year=$year+1900;
GetOptions ('verbose' =>  \$verbose);
#open logfile
if ($verbose){open ($filehandle,'>>',$log);};
if ($verbose){ 
	print $filehandle "------------------\nNew session started at $hour\:$min - $mday $months[$mon]  $year\n-----------------\n";
	print $filehandle "Going to load ini file\n-----------------\n Field \t| Description\n-----------------\n";
};
&load_ini;
if ($verbose){ print $filehandle "\n-----------------\n";};
#Draw window
my $icon = new Win32::GUI::Icon('omc.ico');
my $win_main=Win32::GUI::Window->new(
	-title=>'Outlook Mail Creator',
	-width=>315,
	-height => (40*($itemcount+1)+60),
	-icon => $icon,
	-minimizebox => 0,
	
	
	);
$win_main->SetIcon($icon);
#Add readed params from ini file to window
if ($verbose){
	print $filehandle "Going to create window and it's content\n-----------------\n";
	print $filehandle "Total fields $itemcount\n";
};
my $index=0;
foreach $tmp ( sort keys %inidata) {
	$win_main->AddLabel(-top =>0+(40*$index),-text => "$tmp",-align =>"center" );
	if ($verbose){ print $filehandle "Going to create window field $tmp and it's content is $inidata{$tmp}\n";};
	$label=$win_main->AddTextfield(-align => 'center',-name => "$tmp",-size => [300,20],-pos =>[0,20+(40*$index),]);#-align=> 'left',#-prompt => $inidata{$tmp},
	$label->Append($inidata{$tmp});
	$index=$index+1;
};

$win_main->AddButton(-align => 'center',-ok => 1,-pos =>[100,0+(40*$itemcount)],-name=>'Button1',-text=>"Generate Emails",);
$win_main->Center();
if ($verbose){
	print $filehandle "Window creaated\n-----------------\n";
	print $filehandle "Going to show window\n-----------------\n";
};

$win_main->Show();
Win32::GUI::Dialog();
if ($verbose){ print $filehandle "main Exit\n-----------------\n";};



#subs
sub Button1_Click { 
	&readtmpllist;
	
	&generate_emails;
	}

sub Main_Terminate {
	if ($verbose){ print $filehandle "Main Terminate Exit\n-----------------\n";};
	if ($verbose){ close ($filehandle);};
        -1;
    }
sub readtmpllist {
	
}

sub generate_emails {
	if ($verbose){ print $filehandle "generate email Exit\n-----------------\n";
	close ($filehandle);
	};
	-1;
}


sub load_ini {
	my $input_f;
	if (! open (F_F,'<',"$ini_file")){
		return 0;
	} else {
		foreach $input_f( <F_F> ) {
			my($name_,$descr_) = split /\t/,$input_f;
			chomp($name_);
			chomp($descr_);
			$inidata{$name_}= $descr_;
			if ($verbose){ print $filehandle "$name_\t| $descr_\n";};
			#print $filehandle "Set in hash $name_ \= $inidata{$name_}\n";
			$itemcount++;
			};
	close(F_F);
};
};
