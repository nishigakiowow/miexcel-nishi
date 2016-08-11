#!/bin/bash -xv
echo   "/tmp/log.$(basename $0)" &> /dev/null
exec 2> /tmp/log.$(basename $0)


tmp=/tmp/$$
tmp=/tmp/owow
rm -f $tmp-*

NULL_MOJI="_"
####################
# 文字列情報抽出
####################
cat ../aaa/xl/sharedStrings.xml  |
xmllint  --format - |
awk '{if($1~/^<t>/){sub(/^<t>/,"",$1);sub(/<\/t>$/,"",$NF);print}}' |
awk '{print "s",NR-1,$0}' |
LANG=C sort > $tmp-no_str
# 1:番号 2:文字列
owow=$(awk 'BEGIN{print systime()}')

####################
# 元データ取得
####################
cat .././aaa/xl/worksheets/sheet1.xml  |
xmllint --format - |
tee $tmp-sheet1.xml |
awk 'BEGIN{strcnt=0}
     {if($1=="<c"){cell=$2;type=$3;
        if($0~/t="s"/){str="s";strcnt++}
        else          {str="_"};
      }else if($1~/<v>/){#sub(/^ *<v>/,"",$0);
                        #sub(/<\/v>$/,"",$0);
                        print cell,type,str,$0
                        };
      }' |
awk '{sub(/^[a-z]="/,"",$1);sub(/".*$/,"",$1);
      sub(/^[a-z]="/,"",$2);sub(/".*$/,"",$2);
      sub(/^ *<v>/,"",$4);sub(/<\/v>$/,"",$NF);
      print}' |
LANG=C sort -k3,4 |
join2 +"_" key=3/4 $tmp-no_str |
# 1:セル 2:書式 3:文字だったらs 4:数値 5:文字
awk '{outp=($5=="_")?$4:$5;
      print sprintf("%09d",substr($1,2,1)),sprintf("%03s",substr($1,1,1)),$1,$2,outp,$3}' |
LANG=C sort > $tmp-moto
# 1:９桁R1        2:３桁C1 3:セル 4:書式 5:値
# 6:文字だったらs
owow=$(awk 'BEGIN{print systime()}')


####################
# 更新対象データ
####################
#cat ./data4 |
#cat ./data1000-1 |
#cat ./data10 |
cat .././data100 |
.././xyxy A4 |
tee $tmp-saki |
awk '{print sprintf("%09d",$1),sprintf("%03s",$2),$0}' |
awk '{if($NF~/^[0-9][0-9.]*$/){isnum="_"}  # 文字でない場合
      else                    {isnum="s"}; # 文字の場合
      print $1,$2,$4$3,"1",$5,isnum}' |
tee $tmp-aaa |
#LANG=C sort |
up3 key=1/2 $tmp-moto - |
getlast 1 2 |
# 1:９桁R1        2:３桁C1 3:セル 4:書式 5:値
# 6:文字だったらs
# 文字列対象であれば、シーケンス番号を入れる
awk 'BEGIN{count=0}
     {if($6=="s"){$6=count;count++;}
      print}' > $tmp-newdata
# 1:９桁R1        2:３桁C1 3:セル 4:書式 5:値
# 6:文字カウント
owow=$(awk 'BEGIN{print systime()}')

########################
# 文字列ファイルの作成
########################
awk '$6!="_"{print $5}' $tmp-newdata > $tmp-shredStrings.moto
str_gyo=$(gyo $tmp-shredStrings.moto)
str_uniq_gyo=$(LANG=C sort -u $tmp-shredStrings.moto| gyo)

# はめるテンプレート
cat << FIN > $tmp-sharedStrings.xml.header
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${str_gyo}" uniqueCount="${str_uniq_gyo}">
FIN
cat << FIN | mojihame -l - $tmp-shredStrings.moto > $tmp-sharedStrings.xml.meisai  # 文字でない場合
  <si> <t>%1</t> <phoneticPr fontId="1"/> </si>
FIN
cat << FIN > $tmp-sharedStrings.xml.footer
</sst>
FIN
cat $tmp-sharedStrings.xml.header $tmp-sharedStrings.xml.meisai $tmp-sharedStrings.xml.footer > $tmp-sharedStrings.xml
owow=$(awk 'BEGIN{print systime()}')

########################
# 本体ファイルの作成
########################
# ここまでくる時にはめるのに必要な情報がかけてるかも
awk '{outp=($6=="_")?$5:$6;
      print $1,$2,outp}' $tmp-newdata > $tmp-worksheet.data
awk '/<?xml /,/<sheetData>/' $tmp-sheet1.xml |sed 's/ref="B2:D8"/ref="B2:L30"/' > $tmp-worksheets.header
awk '/<\/sheetData>/,/\/worksheet>/' $tmp-sheet1.xml > $tmp-worksheets.footer

# owowあえてsortしていない
cat $tmp-newdata |
# 1:９桁R1        2:３桁C1 3:セル 4:書式 5:値
# 6:文字カウント
awk 'BEGIN{juni=0}
     {if($1!=maekey){maekey=$1;juni=0}
      else{juni++;};
      if($6=="_"){strflg="_";outp=$5;}
      else{strflg="t=\"s\"";outp=$6;};
      shoshiki=($4=="_")?"_":("s=\""$4"\"");
      print sprintf("%d",$1),$2,sprintf("%06d",juni),$3,shoshiki,strflg,outp}' > $tmp-worksheet.mojihame
owow=$(awk 'BEGIN{print systime()}')
# 1:９桁R1            2:３桁C1 3:行番号 4:セルA1形式 5:書式
# 6:文字列参照かt="s" 7:値    
#cat << FIN > $tmp-template.worksheets.meisai.new
#LABEL-1
#<row r="%1" spans="2:4" x14ac:dyDescent="0.15">
#LABEL-2
#<c r="%4" %5 %6><v>%7</v></c>
#LABEL-2
#</row>
#LABEL-1
#FIN
#mojihame -hLABEL -d_ $tmp-template.worksheets.meisai.new $tmp-worksheet.mojihame  > $tmp-worksheets.meisai
awk 'BEGIN{mae="start"}
     {if($4=="'"${NULL_MOJI}"'"){$4="";};
      if($5=="'"${NULL_MOJI}"'"){$5="";};
      if($6=="'"${NULL_MOJI}"'"){$6="";};
      if($7=="'"${NULL_MOJI}"'"){$7="";};
      if($1!=mae && mae!="start"){print "</row>";};
                  orint "<row r=\""$1"\" spans=\"2:4\" x14ac:dyDescent=\"0.15\">";
                  row_footer="</row>";};
      print "<c r=\""$4"\" "$5" "$6"><v>"$7"</v></c>";
     }
     END{print "</row>"}' $tmp-worksheet.mojihame > $tmp-worksheets.meisai

owow=$(awk 'BEGIN{print systime()}')

cat $tmp-worksheets.header $tmp-worksheets.meisai $tmp-worksheets.footer > $tmp-worksheets.xml

rm -rf ../bbb
rm -f ../bbb.xlsx
cp -pr ../aaa ../bbb
cp -p $tmp-sharedStrings.xml ../bbb/xl/sharedStrings.xml
cp -p $tmp-worksheets.xml ../bbb/xl/worksheets/sheet1.xml
cd ../bbb/ && zip -r ../bbb.zip *
mv ../bbb.zip ../bbb.xlsx
owow=$(awk 'BEGIN{print systime()}')

exit 0
