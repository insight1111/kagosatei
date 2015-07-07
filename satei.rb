#encoding: utf-8
require "win32ole"

def reject_row(value)
	return true if value==42605.0 or value=='診療年月'.encode('cp932') or value==15.0 or value==19.0 or value==51.0 or value==80.0 or value==90.0
	false
end

def msg(msg)
	wsh = WIN32OLE.new('WScript.Shell')
	wsh.Popup(msg)
end

book=WIN32OLE.connect("excel.application").activeworkbook
book.worksheets.add(after: book.sheets(1))
sh=book.sheets(1)
sh3=book.sheets(2)
e=sh.range("A1").end(-4121).row
data=(2..e).select do |i|
	!(reject_row(sh.cells(i,22).value))
end.map do |x| 
	[2,5,14,18,19,20,22,23,24,25,26,27,28].map do |x2| 
		sh.cells(x,x2).value
	end
end
.each.with_index(1) do |x,index| 
	x.each.with_index(1) do |d,i| 
		sh3.cells(index,i).value=d
	end
end
msg("出力完了しました")