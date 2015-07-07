# encoding: utf-8

require "win32ole"
require "byebug"

ex=WIN32OLE.new('excel.application')
book=ex.workbooks.open("C:/projects/kagosatei/H2604国保医科.csv")
sh=book.sheets(1)
addbook=ex.workbooks.add
addsh=addbook.sheets(1)
addsh.name="国保医科"
# addsh=book.sheets(2)
%w(診療月 入外 患者番号 患者氏名 診療科 検査項目 事由 増減 請求内容 補正内容).each.with_index(1) do |x,col|
  addsh.cells(1,col).value=x
end
current=2
last_line=sh.usedrange.rows.count
shinryoukas=(1..last_line).inject([]) do |m,line|
  next m if !(sh.cells(line,8).value =~ Regexp.new(".*科".encode('cp932')))
  m << [sh.cells(line,8).value,line]
  m
end
shinryoukas2=shinryoukas.dup.push([nil,last_line+1])
shinryoukas.map.with_index(0) do |shinryouka,i|
  puts "***#{shinryouka[0]}***".encode('utf-8')
  names=(shinryouka[1]+2..shinryoukas2[i+1][1]-1).inject([]) do |m,line|
    m << [sh.cells(line,18).value, line] if sh.cells(line,18).value
    m
  end
  names2=names.dup
  names2 << [nil,shinryoukas2[i+1][1]-1]
  names.each.with_index(0) do |name,i|
    temp=[]
    group=(name[1]..names2[i+1][1]-1).inject([]) do |m,line|
      next m if sh.cells(line,26).value==nil || sh.cells(line,21).value=="合計".encode("cp932") || sh.cells(line,23).value==nil
      m << line
    end
    p group
    seikyutemp=[]
    _seikyu=[]
    group.map.with_index(0) do |line,j|
      if group[j+1]==nil
        endline=names2[i+1][1]-1
      else
        endline=group[j+1]-1
      end
      seikyu=(line..endline).inject("") do |m,l|
        m=m+(sh.cells(l,26).value+"\n" rescue "")
      end.chomp
      hosei=(line..endline).inject("") do |m,l|
        m=m+(sh.cells(l,28).value+"\n" rescue "\n")
      end.chomp.chomp
      jiyu=[]
      (line..endline).map.with_index(0) do |l,i|
        jiyu << [sh.cells(l,24).value,i] if sh.cells(l,24).value
      end
      seikyutemp << {seikyu: seikyu, hosei: hosei, jiyu: jiyu, zougen: sh.cells(group[j],23).value.to_i} # p "seikyu:#{seikyu}, hosei:#{hosei}\n"
    end
    temp={}

    temp[:seikyus]=seikyutemp
    temp[:name]=name[0]
    temp[:shinryouka]=shinryouka[0]
    _seikyu << temp

    top=current
    addsh.cells(current,1).value="4月"
    addsh.cells(current,2).value= #入外
    addsh.cells(current,3).value=sh.cells(current,19).value #患者番号
    addsh.cells(current,4).value=temp[:name]
    addsh.cells(current,5).value=temp[:shinryouka]
    addsh.cells(current,6).value=sh.cells(current,20).value #検査項目
    temp[:seikyus].each do |_shinryouka|
      _jiyu=[]
      _shinryouka[:jiyu].each{|x| _jiyu[x[1]]=x[0]}
      _jiyu=_jiyu.join("\n")
      addsh.cells(current,7).value=_jiyu
      addsh.cells(current,8).value=_shinryouka[:zougen]
      addsh.cells(current,9).value=_shinryouka[:seikyu]
      addsh.cells(current,10).value=_shinryouka[:hosei]
      # byebug if current==3
      current+=1
    end
    if temp[:seikyus].size > 1
      addsh.range(addsh.cells(top,1),addsh.cells(current-1,1)).merge
      addsh.range(addsh.cells(top,2),addsh.cells(current-1,2)).merge
    end
    p _seikyu
  end #names
end
addsh.columns("A:A").entirecolumn.autofit
addsh.columns("E:F").columnwidth=55
addsh.columns("E:F").verticalalignment=-4160
addsh.columns("C:C").verticalalignment=-4160
addsh.columns("D:D").verticalalignment=-4107
addsh.cells.entirerow.autofit
addbook.saveAs filename: 'c:\temp\H2604集計結果.xlsx'.encode('cp932')
addbook.close
book.close
ex.quit
