#!/usr/bin/perl -w
use Tk;
use DBI;
use utf8;
use Encode;
use Spreadsheet::WriteExcel;

my $dbh=DBI->connect("dbi:SQLite:dbname=money","","") or die $DBI::errstr;

#Fonts
my @Font_entrys=['times','18','normal'];
my @Font_labels=['times','12','bold'];
my @Font_buttons=['times','14','normal'];
#MainWindow
my $mw=MainWindow->new;
$mw->resizable('0','0');
$mw->title("Money");
my $InFrame=$mw->Frame(-height=>200,-borderwidth=>2,-relief => 'groove');
$InFrame->pack(-side=>"top",-fill=>"both");

#Entry DATE NOW
my $EntryDate=$InFrame->Entry(-width=>'10',-font=>@Font_entrys,-bg=>"white");
my ($D, $M, $Y) = (localtime)[3,4,5];	$Y+=1900;	$M++;
if ($D<10) {$D='0'.$D};if ($M<10) {$M='0'.$M};
$EntryDate->insert(0,"$D-$M-$Y");
my $label1=$InFrame->Label(-text=>"Дата",-font=>@Font_labels);
$EntryDate->grid(-row=>0,-column=>0,-padx=>'4');
$label1->grid(-row=>1,-column=>0);




my $EntryRashod=$InFrame->Entry(-width=>'30',-font=>@Font_entrys,-bg=>"white");
my $label2=$InFrame->Label(-text=>"Откуда-Куда",-font=>@Font_labels);
$EntryRashod->grid(-row=>0,-column=>1,-padx=>'10');
$label2->grid(-row=>1,-column=>1);

my $EntrySumma=$InFrame->Entry(-width=>'8',-font=>@Font_entrys,-bg=>"white");
my $label3=$InFrame->Label(-text=>"Сумма",-font=>@Font_labels);
$EntrySumma->grid(-row=>0,-column=>2,-padx=>'10');
$label3->grid(-row=>1,-column=>2);

my $Button_insert_db=$InFrame->Button(-bg=>"#E9E9E9",-activebackground=>"#E1F2F8",-relief=>"ridge",-text=>"Записать",-font=>@Font_buttons,-command=>\&EnterResult);
$Button_insert_db->grid(-row=>0,-column=>3,-rowspan=>2,-padx=>'4');

my $OutFrame=$mw->Frame(-height=>250);
$OutFrame->pack(-side=>"bottom",-fill=>"both");
my $FrameText=$OutFrame->Frame();
my $textout_db=$FrameText->Listbox(-font=>['DejaVu Sans Mono','14','normal'],-setgrid=>'true',
-height=>14,-width=>50);
#Обработка кнопки мыши по листу
$textout_db->bind('<Button-1>',sub {
	my $text=encode('utf8',$textout_db->get($textout_db->curselection()));
	my @ro=(split /:/,$text);
	foreach(@ro){	$_=~ s/\s+$//;	$_=~ s/^\s+//;	};
	$EntryRashod->delete('0','end');
	$EntryRashod->insert('0',decode('utf8',$ro[2]));
	$EntryDate->delete('0','end');
	$EntryDate->insert('0',$ro[1]);
	$EntrySumma->delete('0','end');
	$EntrySumma->insert('0',$ro[3]);
	});

#Удаление записи в listbox Ctrl-D
$mw->bind('<Control-KeyPress-d>'=>sub{
	my $sth;
	if ($textout_db->curselection()>0){
	my $text=encode('utf8',$textout_db->get($textout_db->curselection()));
	my @ro=(split /:/,$text);
	foreach(@ro){	$_=~ s/\s+$//;	$_=~ s/^\s+//;	};
	$sth=$dbh->prepare("delete from money where date=? and rashod=? and summa=?");
	if ($mw->messageBox(-message=>"Вы точно хотите удаоить запись?",-type=>"yesno") eq 'Yes'){
    $ro[1]=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;
    $sth->execute($ro[1],decode('utf8',$ro[2]),$ro[3]);
    $mw->messageBox(-message=>"Запись удалена",-type=>"ok");
    }
};
	
	&Result;	
	});
#окно поиска
$mw->bind('<Control-KeyPress-f>'=>sub{
my $text=encode('utf8',$textout_db->get($textout_db->curselection()));
my @ro=(split /:/,$text);	
foreach(@ro){	$_=~ s/\s+$//;	$_=~ s/^\s+//;	};

my $FindFrame=$mw->Toplevel;
$FindFrame->title("Find");	
my $FrameTopLevel=$FindFrame->Frame()->pack(-side=>'top');
my $FindEntrys=$FrameTopLevel->Entry(-font=>@Font_entrys,-bg=>"white")->pack(-side=>'left');

	if ($ro[2]) {
		$FindEntrys->delete('0','end');
		$FindEntrys->insert('0',decode('utf8',$ro[2]));
	}

my $FindLabel1=$FrameTopLevel->Label(-text=>'c')->pack(-side=>'left');
my $FindFromDate=$FrameTopLevel->Entry(-font=>@Font_entrys,-width=>'10',-bg=>"white")->pack(-side=>'left');
my $FindLabel2=$FrameTopLevel->Label(-text=>'по')->pack(-side=>'left');
my $FindToDate=$FrameTopLevel->Entry(-font=>@Font_entrys,-width=>'10',-bg=>"white")->pack(-side=>'right');
my ($D, $M, $Y) = (localtime)[3,4,5];	$Y+=1900;	$M++;
if ($D<10) {$D='0'.$D};if ($M<10) {$M='0'.$M};
$FindToDate->insert(0,"$D-$M-$Y");
$Y--;
$FindFromDate->insert(0,"$D-$M-$Y");


my $FrameListboxFind=$FindFrame->Frame(-width=>'250')->pack(-side=>'top');
my $FindListbox=$FrameListboxFind->Listbox(-font=>['DejaVu Sans Mono','14','normal'],-setgrid=>'true',
-height=>14,-width=>50)->pack(-side=>'left');	
my $ButtonCloseTopLevel=$FindFrame->Button(-bg=>"#E9E9E9",-activebackground=>"#E1F2F8",-font=>@Font_buttons,-relief=>"ridge",-text=>'Exit',-command=>[$FindFrame=>'destroy'])->pack(-side=>'bottom');

my $FindLabelFrame=$FindFrame->Frame()->pack(-side=>'right');
my $FindLabel=$FindLabelFrame->Label(-text=>'0',-font=>@Font_labels)->pack(-side=>'right',-padx=>'50');
$FindEntrys->bind('<KeyPress>'=>sub{
	my $sth=$dbh->prepare("select distinct * from money where rashod like \"%".$FindEntrys->get."%\" order by date desc");
	unless ($FindEntrys->get) {return 0};
	$sth->execute();
	my $i=1;
	my $summa=0;
	my $Df=$FindFromDate->get;
	$Df=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;
	my $Dt=$FindToDate->get;
	$Dt=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;

	$FindListbox->delete('0.0','end');
	while(my @rows=$sth->fetchrow_array){
		chomp(@rows);
		#~ $rows[0]=~s/(\d\d\d\d)-(\d\d)-(\d\d)/$3-$2-$1/;
		#~ print $rows[0]."\n";
		if (($rows[0] ge $Df)and($rows[0] le $Dt)){
		my $string_db=sprintf( "%2u : %10s : %-15s : %8.2f ",$i,$rows[0],decode("utf8",$rows[1]),$rows[2]);
		$FindListbox->insert("end",$string_db."");
		$i=$i+1;
		$summa+=$rows[2];	
		if ($rows[2]<0) {$FindListbox->itemconfigure("end",-background=>"#FF8B8B");};
		};
    };
    $FindLabel->configure(-text=>$summa);
});
});
	


$FrameText->grid(-row=>0,-column=>0,-rowspan=>10,-padx=>10,-pady=>10,-sticky=>'nw');

$textout_db->grid(-row=>0,-column=>0);

$FrameDate=$OutFrame->Frame(-borderwidth=>2,-relief => 'groove');
$EntryFromDate=$FrameDate->Entry(-width=>10,-text=>"01-$M-$Y",-font=>@Font_entrys,-bg=>"white");
$LabelFromDate=$FrameDate->Label(-text=>'c',-font=>@Font_labels);
$EntryToDate=$FrameDate->Entry(-width=>10,-text=>"$D-$M-$Y",-font=>@Font_entrys,-bg=>"white");
$LabelToDate=$FrameDate->Label(-text=>'по',-font=>@Font_labels);

$FrameDate->grid(-row=>0,-column=>1,-padx=>10,-pady=>10,-sticky=>'n');
$EntryFromDate->grid(-row=>0,-column=>1,-padx=>10,-pady=>5,-sticky=>'w');
$LabelFromDate->grid(-row=>0,-column=>0,-padx=>5,-pady=>5,-sticky=>'e');

$EntryToDate->grid(-row=>1,-column=>1,-padx=>10,-pady=>5,-sticky=>'w');
$LabelToDate->grid(-row=>1,-column=>0,-padx=>5,-pady=>5,-sticky=>'e');

$ButtonResult=$OutFrame->Button(-bg=>"#E9E9E9",-activebackground=>"#E1F2F8",-relief=>"ridge",-text=>'Вывести на экран',-font=>@Font_buttons,-width=>17,-command=>\&Result);
$ButtonResult->grid(-row=>1,-column=>1);
$ButtonResult1=$OutFrame->Button(-bg=>"#E9E9E9",-activebackground=>"#E1F2F8",-relief=>"ridge",-text=>'Вывести в Excel',-font=>@Font_buttons,-width=>17,
-command=>\&Result1);
$ButtonResult1->grid(-row=>2,-column=>1);

$FrameSumma=$OutFrame->Frame(-bg=>'light blue');
my ($LabelVarPlus,$LabelVarMinus,$LabelVarSumma,$LabelVarSummaYear);
$LabelPlus=$FrameSumma->Label(-textvariable=>\$LabelVarPlus,-bg=>'light blue',-font=>['times','14','bold']);
$LabelMinus=$FrameSumma->Label(-textvariable=>\$LabelVarMinus,-fg=>'red',-bg=>'light blue',-font=>['times','14','bold']);
$LabelSumma=$FrameSumma->Label(-textvariable=>\$LabelVarSumma,-fg=>'dark green',-bg=>'light blue',-font=>['times','14','bold']);
$LabelSummaYear=$FrameSumma->Label(-textvariable=>\$LabelVarSummaYear,-fg=>'maroon',-bg=>'light blue',-font=>['times','14','bold']);
$FrameSumma->grid(-row=>3,-column=>1,-sticky=>'n');
$LabelPlus->grid(-row=>0,-column=>0,-padx=>5,-pady=>3,-sticky=>'e');
$LabelMinus->grid(-row=>1,-column=>0,-padx=>5,-pady=>3,-sticky=>'e');
$LabelSumma->grid(-row=>2,-column=>0,-padx=>5,-pady=>3,-sticky=>'e');
$LabelSummaYear->grid(-row=>3,-column=>0,-padx=>5,-pady=>3,-sticky=>'e');
&Result;



MainLoop();

#Нажатие кнопки запись
sub EnterResult{	
if ($textout_db->curselection()>"0"){
my $text=encode('utf8',$textout_db->get($textout_db->curselection()));
my @ro=(split /:/,$text);
foreach(@ro){	$_=~ s/\s+$//;	$_=~ s/^\s+//;	};	
$ro[1]=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;
my $Df=$EntryDate->get;
$Df=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;

my $sth=$dbh->prepare("update money set date=?, rashod=?, summa=? where date=? and rashod=? and summa=?");
$sth->execute($Df,$EntryRashod->get,$EntrySumma->get,$ro[1],decode('utf8',$ro[2]),$ro[3]);
$mw->messageBox(-message=>"Изменения прошли успешно",-type=>"ok");	
	} else{
my $sth=$dbh->prepare("insert into money (date,rashod,summa) values(?,?,?)");
my $Df=$EntryDate->get;
$Df=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;
$sth->execute($Df,$EntryRashod->get,$EntrySumma->get);
$mw->messageBox(-message=>"Запись прошла успешно",-type=>"ok") if defined($sth);	
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
$textout_db->delete('0.0','end');

while(my @rows=$sth->fetchrow_array){
	chomp(@rows);
	$sumall+=$rows[2];
	if (($rows[0] ge $Df)and($rows[0] le $Dt)){
	$rows[0]=~s/(\d\d\d\d)-(\d\d)-(\d\d)/$3-$2-$1/;	
	my $string_db=sprintf( "%2u : %10s : %-15s : %8.2f ",$i,$rows[0],decode("utf8",$rows[1]),$rows[2]);
    $textout_db->insert("end",$string_db."");
    $i=$i+1;	
    if ($rows[2]>0){$sump+=$rows[2];} else {$summ+=$rows[2];$textout_db->itemconfigure("end",-background=>"#FF8B8B");};
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
