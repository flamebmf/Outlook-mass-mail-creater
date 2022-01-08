#perl2exe_info FileDescription =Outlook_Mass_Mail_Creator
#perl2exe_info FileVersion=1.0
#perl2exe_info InternalName=Outlook_Mail_Creator.pl
#perl2exe_info LegalCopyright=Akulich Dmitry akulich.d@gmail.com
#perl2exe_info ProductName=Outlook Mail Creator
#perl2exe_info ProductVersion=1
use Getopt::Long;
use Cwd qw(cwd);
use strict;
use warnings;
use Config::Tiny;
use Win32::GUI();
use Win32::OLE();
#use Win32::OLE::Const 'Microsoft Outlook';
#files
my $config = Config::Tiny->read( "omc.ini", 'utf8' );
my $ini_file=$config->{files}{fields}; # file with fields and it's describtion
my $emailfileslist=$config->{files}{emailtemlates};# list of filenames which include mail template
my $inputdata=$config->{files}{inputdata}; # file with data which must be loaded in fields
my $adddb_f=$config->{files}{adddb};
my $log=$config->{files}{log}; # verbose log
#hashes
my %emailtmpl;
my %storedb;
my %inidata;
my %labels;
# vars
my $filehandle;
my $tmp;
my $label;
my $itemcount=0;
my $verbose=$config->{params}{verbose};
my @months = qw( Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec );
my @days = qw(Sun Mon Tue Wed Thu Fri Sat Sun);
#main

if (!$verbose) {GetOptions ('verbose' =>  \$verbose)};
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
	-width=>316,
	-height => (40*($itemcount+1)+40),
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
foreach $tmp ( keys %inidata) {
	$win_main->AddLabel(-top =>0+(40*$index),-text => "$tmp",-align =>"center" );
	if ($verbose){ print $filehandle "Going to create text field $tmp and it's content is $inidata{$tmp}\n";};
	$labels{$tmp}=$win_main->AddTextfield(-align => 'center',-name => "$tmp",-size => [300,20],-pos =>[3,20+(40*$index),]);#-align=> 'left',#-prompt => $inidata{$tmp},
	$labels{$tmp}->Append($inidata{$tmp});
	$index=$index+1;
};
$win_main->AddButton(-align => 'center',-ok => 0,-pos =>[28,0+(40*$itemcount)],-size => [100,20],-name=>'Button2',-text=>"Load Data",);
$win_main->AddButton(-align => 'center',-ok => 1,-pos =>[178,0+(40*$itemcount)],-size => [100,20],-name=>'Button1',-text=>"Generate Emails",);
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
	my $to;
	my $cc;
	my $subject;
	my $body="";
	my $tmp;
	my %input_h;
	
	if ($verbose){ print $filehandle "Going to read fields\n";};
	foreach $tmp (sort keys %labels){
		$input_h{$tmp}=$labels{$tmp}->GetLine(0);
		if ($verbose){ print $filehandle "readed field $tmp data $input_h{$tmp}\n";};
		}		
	if ($verbose){ print $filehandle "\n-----------------\nGoing to read templates\n-----------------\n";};
	&readtmpllist();
	my $ol = Win32::OLE->GetActiveObject('Outlook.Application') || Win32::OLE->new('Outlook.Application', 'Quit');
	foreach $tmp ( sort keys %emailtmpl) {
		$body="";
	if (! open (T_F,'<',"$emailtmpl{$tmp}")){ 
		return 0;
		} else {
			if ($verbose){ print $filehandle "Going to filter template $tmp\n";};
			$to = <T_F>;
			$cc = <T_F>;
			$subject = <T_F>;
			foreach my $line (<T_F>) {
				$body=$body . $line;
			};
			close (T_F);
			
		};
	if ($verbose){ print $filehandle "Going to filter template $tmp\n";};
	foreach $tmp (sort keys %input_h){ #filter strings
		$to=~s/$tmp/$input_h{$tmp}/g;
		$cc=~s/$tmp/$input_h{$tmp}/g;
		$subject=~s/$tmp/$input_h{$tmp}/g;
		$body=~s/$tmp/$input_h{$tmp}/g;
		if ($verbose){ print $filehandle "filtering field $tmp data $input_h{$tmp}\n";};
		}			
	my $email=$ol->CreateItem(0);
	$email->{'To'}= $to;
	$email->{'CC'}= $cc;
	$email->{'Subject'}= $subject;
	$email->{'BodyFormat'} = 'olFormatHTML';
	$email->{'HTMLBody'} = $body;
	$email->save();
	$email->Display();
	};
    $ol->{Visible} = 1;
    

    #&generate_emails;
	}

sub Button2_Click { 
	my $input_f;
	my %input_h;
	if ($verbose){ print $filehandle "Going to load data\n-----------------\n";};
	if (! open (F_F,'<',"$inputdata")){
		return 0;
	} else {
		foreach $input_f( <F_F> ) {
			my($name_,$descr_) = split /\t/,$input_f;
			chomp($name_);
			chomp($descr_);
			$input_h{$name_}= $descr_;
			if ($verbose){ print $filehandle "$name_\t| $descr_\n";};
			};
	close(F_F);
	foreach $tmp (sort keys %labels){
		if ($verbose){ print $filehandle "Going to load data $input_h{$tmp} to text field $tmp \n";};
		$labels{$tmp}->SelectAll();
		$labels{$tmp}->Clear();
		$labels{$tmp}->Append($input_h{$tmp});
		}
	};
};
sub Main_Terminate {
	if ($verbose){ print $filehandle "Main Terminate Exit\n-----------------\n";};
	if ($verbose){ close ($filehandle);};
        -1;
    }
sub readtmpllist {
	my $input_f;
	if (! open (F_F,'<',"$emailfileslist")){
		return 0;
	} else {
		foreach $input_f( <F_F> ) {
			my($name_,$descr_) = split /\t/,$input_f;
			chomp($name_);
			chomp($descr_);
			$emailtmpl{$name_}= $descr_;
			if ($verbose){ print $filehandle "$name_\t| $descr_\n";};
			#print $filehandle "Set in hash $name_ \= $inidata{$name_}\n";
			#$itemcount++;
			};
	close(F_F);
};
	
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
