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
$mw->title("Money");
my $InFrame=$mw->Frame(-height=>200,-borderwidth=>2,-relief => 'groove');
$InFrame->pack(-side=>"top",-fill=>"both");

#Entry DATE NOW
my $EntryDate=$InFrame->Entry(-width=>'10',-font=>@Font_entrys);
my ($D, $M, $Y) = (localtime)[3,4,5];	$Y+=1900;	$M++;
if ($D<10) {$D='0'.$D};
if ($M<10) {$M='0'.$M};
$EntryDate->insert(0,"$D-$M-$Y");
my $label1=$InFrame->Label(-text=>"Дата",-font=>@Font_labels);
$EntryDate->grid(-row=>0,-column=>0,-padx=>'4');
$label1->grid(-row=>1,-column=>0);




my $EntryRashod=$InFrame->Entry(-width=>'30',-font=>@Font_entrys);
my $label2=$InFrame->Label(-text=>"Откуда-Куда",-font=>@Font_labels);
$EntryRashod->grid(-row=>0,-column=>1,-padx=>'10');
$label2->grid(-row=>1,-column=>1);

my $EntrySumma=$InFrame->Entry(-width=>'8',-font=>@Font_entrys);
my $label3=$InFrame->Label(-text=>"Сумма",-font=>@Font_labels);
$EntrySumma->grid(-row=>0,-column=>2,-padx=>'10');
$label3->grid(-row=>1,-column=>2);

my $Button_insert_db=$InFrame->Button(-text=>"Записать",-font=>@Font_buttons,-command=>sub{
my $sth=$dbh->prepare("insert into money (date,rashod,summa) values(?,?,?)");
$sth->execute($EntryDate->get,$EntryRashod->get,$EntrySumma->get);
 $mw->messageBox(-message=>"Запись прошла успешно",-type=>"ok") if defined($sth);	
	
	});
$Button_insert_db->grid(-row=>0,-column=>3,-rowspan=>2,-padx=>'4');

my $OutFrame=$mw->Frame(-height=>250);
$OutFrame->pack(-side=>"bottom",-fill=>"both");
my $FrameText=$OutFrame->Frame();
my $scrollbarv1=$FrameText->Scrollbar();
my $textout_db=$FrameText->Text(-font=>['DejaVu Sans Mono','14','normal'],-setgrid=>'true',
-height=>14,-width=>50,-yscrollcommand=>['set',$scrollbarv1]);
$scrollbarv1->configure(-command=>['yview',$textout_db]);

$FrameText->grid(-row=>0,-column=>0,-rowspan=>10,-padx=>10,-pady=>10,-sticky=>'nw');

$textout_db->grid(-row=>0,-column=>0);
$scrollbarv1->grid(-row=>0,-column=>1,-sticky=>'ns');

$FrameDate=$OutFrame->Frame(-borderwidth=>2,-relief => 'groove');
$EntryFromDate=$FrameDate->Entry(-width=>10,-text=>"01-$M-$Y",-font=>@Font_entrys);
$LabelFromDate=$FrameDate->Label(-text=>'c',-font=>@Font_labels);
$EntryToDate=$FrameDate->Entry(-width=>10,-text=>"$D-$M-$Y",-font=>@Font_entrys);
$LabelToDate=$FrameDate->Label(-text=>'по',-font=>@Font_labels);

$FrameDate->grid(-row=>0,-column=>1,-padx=>10,-pady=>10,-sticky=>'n');
$EntryFromDate->grid(-row=>0,-column=>1,-padx=>10,-pady=>5,-sticky=>'w');
$LabelFromDate->grid(-row=>0,-column=>0,-padx=>5,-pady=>5,-sticky=>'e');

$EntryToDate->grid(-row=>1,-column=>1,-padx=>10,-pady=>5,-sticky=>'w');
$LabelToDate->grid(-row=>1,-column=>0,-padx=>5,-pady=>5,-sticky=>'e');

$ButtonResult=$OutFrame->Button(-text=>'Вывести на экран',-font=>@Font_buttons,-width=>17,-command=>\&Result);
$ButtonResult->grid(-row=>1,-column=>1);
$ButtonResult1=$OutFrame->Button(-text=>'Вывести в Excel',-font=>@Font_buttons,-width=>17,
-command=>\&Result1);
$ButtonResult1->grid(-row=>2,-column=>1);

$FrameSumma=$OutFrame->Frame(-bg=>'light blue');
my ($LabelVarPlus,$LabelVarMinus,$LabelVarSumma);
$LabelPlus=$FrameSumma->Label(-textvariable=>\$LabelVarPlus,-bg=>'light blue',-font=>['times','14','bold']);
$LabelMinus=$FrameSumma->Label(-textvariable=>\$LabelVarMinus,-fg=>'red',-bg=>'light blue',-font=>['times','14','bold']);
$LabelSumma=$FrameSumma->Label(-textvariable=>\$LabelVarSumma,-fg=>'dark green',-bg=>'light blue',-font=>['times','14','bold']);
$FrameSumma->grid(-row=>3,-column=>1,-sticky=>'n');
$LabelPlus->grid(-row=>0,-column=>0,-padx=>5,-pady=>3,-sticky=>'e');
$LabelMinus->grid(-row=>1,-column=>0,-padx=>5,-pady=>3,-sticky=>'e');
$LabelSumma->grid(-row=>2,-column=>0,-padx=>5,-pady=>3,-sticky=>'e');
&Result;



MainLoop();







#SELECT FROM DB ALL TO SCREEN
sub Result{
my $sth=$dbh->prepare("select * from money order by date desc");
$sth->execute();
my $i=1;
my $Df=$EntryFromDate->get;
$Df=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;
my $Dt=$EntryToDate->get;
$Dt=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;
my $sump=0;
my $summ=0;
$textout_db->delete('1.0','end');

while(my @rows=$sth->fetchrow_array){
	chomp(@rows);
	my $Dn=$rows[0];
	$Dn=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;
	if (($Dn ge $Df)and($Dn le $Dt)){
	my $string_db=sprintf( "%2u | %10s | %-15s | %8.2f ",$i,$rows[0],decode("utf8",$rows[1]),$rows[2]);
    $textout_db->insert('end',$string_db."\n");
    
    $i=$i+1;	
    if ($rows[2]>0){$sump+=$rows[2];} else {$summ+=$rows[2];};
    }
}
$LabelVarPlus="+$sump";
$LabelVarMinus="$summ";
$sump=$sump+$summ;
$LabelVarSumma="Всего $sump";
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
	my $Dn=$rows[0];
	$Dn=~s/(\d\d)-(\d\d)-(\d\d\d\d)/$3-$2-$1/;
	if (($Dn ge $Df)and($Dn le $Dt)){
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
