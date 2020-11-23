#!/usr/bin/perl -w
use Tkx;
use DBI;
use utf8;
use Encode;
use Spreadsheet::WriteExcel;

my $dbh=DBI->connect("dbi:SQLite:dbname=money","","") or die $DBI::errstr;

#Fonts
my @Font_entrys=['PT Mono','18','normal'];
my @Font_labels=['PT Mono','12','bold'];
my @Font_buttons=['PT Mono','14','normal'];
#MainWindow
my $mw=Tkx::widget->new(".");
$mw->g_wm_title("Money");
my $InFrame=$mw->new_frame(-height=>200,-borderwidth=>2,-relief => 'groove');
$InFrame->g_pack(-side=>"top",-fill=>"both");

#Entry DATE NOW
my $EntryDate=$InFrame->new_entry(-width=>'10',-font=>@Font_entrys);
my ($D, $M, $Y) = (localtime)[3,4,5];	$Y+=1900;	$M++;
if ($D<10) {$D='0'.$D};if ($M<10) {$M='0'.$M};
$EntryDate->insert(0,"$D-$M-$Y");
my $label1=$InFrame->new_label(-text=>"Дата",-font=>@Font_labels);
$EntryDate->g_grid(-row=>0,-column=>0,-padx=>'4');
$label1->g_grid(-row=>1,-column=>0);
#~ 


my $EntryRashod=$InFrame->new_entry(-width=>'30',-font=>@Font_entrys);
my $label2=$InFrame->new_label(-text=>"Откуда-Куда",-font=>@Font_labels);
$EntryRashod->g_grid(-row=>0,-column=>1,-padx=>'10');
$label2->g_grid(-row=>1,-column=>1);

my $EntrySumma=$InFrame->new_entry(-width=>'8',-font=>@Font_entrys);
my $label3=$InFrame->new_label(-text=>"Сумма",-font=>@Font_labels);
$EntrySumma->g_grid(-row=>0,-column=>2,-padx=>'10');
$label3->g_grid(-row=>1,-column=>2);

my $Button_insert_db=$InFrame->new_button(-text=>"Записать",-font=>@Font_buttons,-command=>\&EnterResult);
$Button_insert_db->g_grid(-row=>0,-column=>3,-rowspan=>2,-padx=>'4');

my $OutFrame=$mw->new_frame(-height=>250);
$OutFrame->g_pack(-side=>"bottom",-fill=>"both");
my $FrameText=$OutFrame->new_frame();
my $scrollbarv1=$FrameText->new_scrollbar();
my $textout_db=$FrameText->new_listbox(-font=>['PT Mono','14','normal'],-setgrid=>'true',
-height=>14,-width=>50,-yscrollcommand=>[$scrollbarv1,'set']);
#Обработка кнопки мыши по листу
Tkx::bind($textout_db,"<Button>",sub {
	my $wer=$textout_db->curselection;
	 if ($wer) {
	my $text=encode('utf8',$textout_db->get($wer));
	my @ro=(split /:/,$text);
	foreach(@ro){	$_=~ s/\s+$//;	$_=~ s/^\s+//;	};
	$EntryRashod->delete('0','end');
	$EntryRashod->insert('0',decode('utf8',$ro[2]));
	#~ $EntryDate->delete('0','end');
	#~ $EntryDate->insert('0',$ro[1]);
	#~ $EntrySumma->delete('0','end');
	#~ $EntrySumma->insert('0',$ro[3]);
	}});

#Удаление записи в listbox Ctrl-D
$mw->g_bind('<Control-KeyPress-d>'=>sub{
	my $sth;
	if ($textout_db->curselection()>0){
	my $text=encode('utf8',$textout_db->get($textout_db->curselection()));
	
	
	#~ my $text=encode('utf8',$textout_db->get($textout_db->curselection()));
	my @ro=(split /:/,$text);
	foreach(@ro){	$_=~ s/\s+$//;	$_=~ s/^\s+//;	};
	$sth=$dbh->prepare("delete from money where date=? and rashod=? and summa=?");
	if (Tkx::tk___messageBox(-message=>"Вы точно хотите удаоить запись?",-type=>"yesno") eq 'Yes'){
    $ro[1]=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;
    $sth->execute($ro[1],decode('utf8',$ro[2]),$ro[3]);
    Tkx::tk___messageBox(-message=>"Запись удалена",-type=>"ok");
    }
};
	
	&Result;	
	});
	
	
$scrollbarv1->configure(-command=>[$textout_db,'yview']);

$FrameText->g_grid(-row=>0,-column=>0,-rowspan=>10,-padx=>10,-pady=>10,-sticky=>'nw');

$textout_db->g_grid(-row=>0,-column=>0);
$scrollbarv1->g_grid(-row=>0,-column=>1,-sticky=>'ns');

$FrameDate=$OutFrame->new_frame(-borderwidth=>2,-relief => 'groove');
$EntryFromDate=$FrameDate->new_entry(-width=>10,-text=>"01",-font=>@Font_entrys);
$LabelFromDate=$FrameDate->new_label(-text=>'c',-font=>@Font_labels);
$EntryFromDate->insert(0,"01-$M-$Y");
$EntryToDate=$FrameDate->new_entry(-width=>10,-font=>@Font_entrys);
$EntryToDate->insert(0,"$D-$M-$Y");
$LabelToDate=$FrameDate->new_label(-text=>'по',-font=>@Font_labels);

$FrameDate->g_grid(-row=>0,-column=>1,-padx=>10,-pady=>10,-sticky=>'n');
$EntryFromDate->g_grid(-row=>0,-column=>1,-padx=>10,-pady=>5,-sticky=>'w');
$LabelFromDate->g_grid(-row=>0,-column=>0,-padx=>5,-pady=>5,-sticky=>'e');

$EntryToDate->g_grid(-row=>1,-column=>1,-padx=>10,-pady=>5,-sticky=>'w');
$LabelToDate->g_grid(-row=>1,-column=>0,-padx=>5,-pady=>5,-sticky=>'e');

$ButtonResult=$OutFrame->new_button(-text=>'Вывести на экран',-font=>@Font_buttons,-width=>17,-command=>\&Result);
$ButtonResult->g_grid(-row=>1,-column=>1);
$ButtonResult1=$OutFrame->new_button(-text=>'Вывести в Excel',-font=>@Font_buttons,-width=>17,
-command=>\&Result1);
$ButtonResult1->g_grid(-row=>2,-column=>1);

$FrameSumma=$OutFrame->new_frame(-bg=>'light blue');
my ($LabelVarPlus,$LabelVarMinus,$LabelVarSumma,$LabelVarSummaYear);
$LabelPlus=$FrameSumma->new_label(-textvariable=>\$LabelVarPlus,-bg=>'light blue',-font=>['times','14','bold']);
$LabelMinus=$FrameSumma->new_label(-textvariable=>\$LabelVarMinus,-fg=>'red',-bg=>'light blue',-font=>['times','14','bold']);
$LabelSumma=$FrameSumma->new_label(-textvariable=>\$LabelVarSumma,-fg=>'dark green',-bg=>'light blue',-font=>['times','14','bold']);
$LabelSummaYear=$FrameSumma->new_label(-textvariable=>\$LabelVarSummaYear,-fg=>'maroon',-bg=>'light blue',-font=>['times','14','bold']);
$FrameSumma->g_grid(-row=>3,-column=>1,-sticky=>'n');
$LabelPlus->g_grid(-row=>0,-column=>0,-padx=>5,-pady=>3,-sticky=>'e');
$LabelMinus->g_grid(-row=>1,-column=>0,-padx=>5,-pady=>3,-sticky=>'e');
$LabelSumma->g_grid(-row=>2,-column=>0,-padx=>5,-pady=>3,-sticky=>'e');
$LabelSummaYear->g_grid(-row=>3,-column=>0,-padx=>5,-pady=>3,-sticky=>'e');
&Result;



Tkx::MainLoop();

#Нажатие кнопки запись
sub EnterResult{	
if ($textout_db->curselection()>0){
my $text=encode('utf8',$textout_db->get($textout_db->curselection()));
my @ro=(split /:/,$text);
foreach(@ro){	$_=~ s/\s+$//;	$_=~ s/^\s+//;	};	
$ro[1]=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;
my $Df=$EntryDate->get;
$Df=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;

my $sth=$dbh->prepare("update money set date=?, rashod=?, summa=? where date=? and rashod=? and summa=?");
$sth->execute($Df,$EntryRashod->get,$EntrySumma->get,$ro[1],decode('utf8',$ro[2]),$ro[3]);
Tkx::tk___messageBox(-message=>"Изменения прошли успешно",-type=>"ok");	
	} else{
my $sth=$dbh->prepare("insert into money (date,rashod,summa) values(?,?,?)");
my $Df=$EntryDate->get;
$Df=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;
$sth->execute($Df,$EntryRashod->get,$EntrySumma->get);
Tkx::tk___messageBox(-message=>"Запись прошла успешно",-type=>"ok") if defined($sth);	
};
&Result;
		}




#SELECT FROM DB ALL TO SCREEN
sub Result{
my $sth=$dbh->prepare("select distinct * from money order by date desc");
$sth->execute();
my $i=1;

my ($D, $M, $Y) = (localtime)[3,4,5];	$Y+=1900;	$M++;
if ($D<10) {$D='0'.$D};
if ($M<10) {$M='0'.$M};
$EntryDate->delete('0','end');
$EntryDate->insert(0,"$D-$M-$Y");

my $Df=$EntryFromDate->get;
$Df=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;
my $Dt=$EntryToDate->get;
$Dt=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;
my $sump=0;
my $summ=0;
my $sumall=0;
$textout_db->delete('0','end');

while(my @rows=$sth->fetchrow_array){
	chomp(@rows);
	$sumall+=$rows[2];
	if (($rows[0] ge $Df)and($rows[0] le $Dt)){
	$rows[0]=~s/(\d\d\d\d)-(\d\d)-(\d\d)/$3-$2-$1/;	
	my $string_db=sprintf( "%2u : %10s : %-15s : %8.2f ",$i,$rows[0],decode("utf8",$rows[1]),$rows[2]);
    $textout_db->insert("end",$string_db."\n");
    $i=$i+1;	
    if ($rows[2]>0){$sump+=$rows[2];} else {$summ+=$rows[2];$textout_db->itemconfigure("end",-background=>"#FF0F00");};
    }
}
$LabelVarPlus="+$sump";
$LabelVarMinus="$summ";
$sump=$sump+$summ;
$LabelVarSumma="Итог за м-ц $sump";
$LabelVarSummaYear="Всего $sumall";
$EntryRashod->delete('0','end');
$EntrySumma->delete('0','end');
}

#SELECT FROM DB ALL TO EXCEL
sub Result1{
my $sth=$dbh->prepare("select * from money order by date desc");
$sth->execute();
my $i=1;
my $Df=$EntryFromDate->get;
$Df=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;
my $Dt=$EntryToDate->get;
$Dt=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;
my $sum=0;
my $workbook = Spreadsheet::WriteExcel->new('Result.xls');
my $worksheet = $workbook->add_worksheet('Report');

while(my @rows=$sth->fetchrow_array){
	chomp(@rows);
	if (($rows[0] ge $Df)and($rows[0] le $Dt)){
    $rows[0]=~s/(\d\d\d\d)-(\d\d)-(\d\d)/$3-$2-$1/;	
	$worksheet->write($i, 1,$rows[0]);
    $worksheet->write($i, 2,decode("utf8",$rows[1]));
    $worksheet->write($i, 3,$rows[2]);
    $i=$i+1;	
    $sum+=$rows[2];
    }
};
$worksheet->write($i, 3,$sum);
$workbook->close();
exec 'Result.xls';
    
}
