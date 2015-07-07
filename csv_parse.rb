# encoding: utf-8

require "win32ole"
require "byebug"

ex=WIN32OLE.new('excel.application')

# bookは査定CSV
book=ex.workbooks.open("C:/projects/kagosatei/H2604国保医科.csv")
sh=book.sheets(1)

# resultBookはパース後のbook
resultBook=ex.workbooks.add
resultSheet=resultBook.sheets(1)
resultSheet.name="国保医科"

# パース後に入れ込むヘッダー
%w(診療月 入外 患者番号 患者氏名 診療科 検査項目 事由 増減 請求内容 補正内容).each.with_index(1) do |x,col|
  resultSheet.cells(1,col).value=x
end

# 集計資料対象は2行目から開始
current=2
last_line=sh.usedrange.rows.count

# 査定CSVから○○科の部分を抽出し、その行数をストアする
# shinryoukas=[[診療科名, その開始行],[]...]
shinryoukas=(1..last_line).inject([]) do |m,line|
  next m if !(sh.cells(line,8).value =~ Regexp.new(".*科".encode('cp932')))
  m << [sh.cells(line,8).value,line]
  m
end

# パースする範囲は診療科表記のあった行から、最終行(lastline)まで。
# 最終行を入れる一時変数として_shinryoukasを作成　
_shinryoukas=shinryoukas.dup.push([nil,last_line+1])
shinryoukas.map.with_index(0) do |shinryouka,i|
  puts "***#{shinryouka[0]}***".encode('utf-8')
  name_col=18

  # 各診療科の範囲内で患者名を抽出。その行数をストアする
  # names=[[患者名, その開始行],...]
  names=(shinryouka[1]+2.._shinryoukas[i+1][1]-1).inject([]) do |m,line|
    m << [sh.cells(line,name_col).value, line] if sh.cells(line,name_col).value
    m
  end

  # 1患者あたりのパース範囲は名前が出てきた行から、次の名前が出てくる行-1
  # 範囲の最終行を保持する一時変数として_namesを作成
  _names=names.dup.push([nil,_shinryoukas[i+1][1]-1])
  names.each.with_index(0) do |name,i|
    temp=[]
    # まず請求内容をグループに分ける。グループの境界は
    # "合計"がなく、請求書内容が空でない行で、点数が記載されている行が区切り
    goukei_col=21
    tensu_col=23
    naiyo_col=26
    group=(name[1].._names[i+1][1]-1).inject([]) do |m,line|
      next m if sh.cells(line,naiyo_col).value==nil || sh.cells(line,goukei_col).value=="合計".encode("cp932") || sh.cells(line,tensu_col).value==nil
      m << line
    end
    p group
    seikyutemp=[]
    _seikyu=[]

    # 各グループごとの請求内容を抽出する
    group.map.with_index(0) do |line,j|
      if group[j+1]==nil
        endline=_names[i+1][1]-1
      else
        endline=group[j+1]-1
      end

      # seikyuは請求内容
      # hoseiは査定で補正された内容
      seikyu=(line..endline).inject("") do |m,l|
        m=m+(sh.cells(l,naiyo_col).value+"\n" rescue "")
      end.chomp
      hosei=(line..endline).inject("") do |m,l|
        m=m+(sh.cells(l,28).value+"\n" rescue "\n")
      end.chomp.chomp
      jiyu=[]
      (line..endline).map.with_index(0) do |l,i|
        jiyu << [sh.cells(l,24).value,i] if sh.cells(l,24).value
      end
      seikyutemp << {seikyu: seikyu, hosei: hosei, jiyu: jiyu, zougen: sh.cells(group[j],tensu_col).value.to_i} # p "seikyu:#{seikyu}, hosei:#{hosei}\n"
    end
    temp={}

    temp[:seikyus]=seikyutemp
    temp[:name]=name[0]
    temp[:shinryouka]=shinryouka[0]
    _seikyu << temp

    top=current
    resultSheet.cells(current,1).value="4月"
    resultSheet.cells(current,2).value= #入外
    resultSheet.cells(current,3).value=sh.cells(current,19).value #患者番号
    resultSheet.cells(current,4).value=temp[:name]
    resultSheet.cells(current,5).value=temp[:shinryouka]
    resultSheet.cells(current,6).value=sh.cells(current,20).value #検査項目
    temp[:seikyus].each do |_shinryouka|
      _jiyu=[]
      _shinryouka[:jiyu].each{|x| _jiyu[x[1]]=x[0]}
      _jiyu=_jiyu.join("\n")
      resultSheet.cells(current,7).value=_jiyu
      resultSheet.cells(current,8).value=_shinryouka[:zougen]
      resultSheet.cells(current,9).value=_shinryouka[:seikyu]
      resultSheet.cells(current,10).value=_shinryouka[:hosei]
      # byebug if current==3
      current+=1
    end

    # 請求内容が複数あるばあいは、セルを結合する
    if temp[:seikyus].size > 1
      resultSheet.range(resultSheet.cells(top,1),resultSheet.cells(current-1,1)).merge
      resultSheet.range(resultSheet.cells(top,2),resultSheet.cells(current-1,2)).merge
    end
    p _seikyu
  end #names
end

# 最後に表を整形する
resultSheet.columns("A:A").entirecolumn.autofit
resultSheet.columns("E:F").columnwidth=55
resultSheet.columns("E:F").verticalalignment=-4160
resultSheet.columns("C:C").verticalalignment=-4160
resultSheet.columns("D:D").verticalalignment=-4107
resultSheet.cells.entirerow.autofit
resultBook.saveAs filename: 'c:\temp\H2604集計結果.xlsx'.encode('cp932')
resultBook.close
book.close
ex.quit
