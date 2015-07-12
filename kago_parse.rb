# encoding: utf-8

# 同一フォルダにあるエクセルファイルから、ピボット用の
# csvファイルを作成する

STDOUT.sync = true
require "win32ole"

# wshを利用したメッセージボックス
def msgbox(msg,title)
  wsh=WIN32OLE.new("WScript.shell")
  wsh.popup(msg,0,title,0+64)
end

basedirname=File.expand_path(File.dirname(__FILE__))

xl=WIN32OLE.new('excel.application')
data=[]
Dir.glob(basedirname+"/*.xls") do |file|
	
	book=xl.workbooks.open(:filename => file, :readonly => true)
	puts book.name.encode('utf-8')
	satei_col=11
	["★入院査定", "■入院過誤", "★外来査定 ", "■外来過誤"].each do |s|
		sh=book.sheets(s.encode('cp932'))
		(2..sh.range("A:A").end(-4121).row).each do |line|
			puts line
			sh.cells(line,satei_col).value.split.each {|v|
				next if v =~ /^\d/
				data << {
					kensa:      v.encode('utf-8').split(/　/)[0],
					shinryouka: sh.cells(line,4).value.encode('utf-8'),
					category:   s[1..-1].encode('utf-8'),
					month:      File.basename(file)[3..4],
					kokuho:     sh.cells(line,1).value.encode('utf-8'),
					shujii:     sh.cells(line,5).value.encode('utf-8'),
					koui:       sh.cells(line,8).value.encode('utf-8'),
					riyu:       sh.cells(line,9).value.encode('utf-8')
				}
			}
		end

	end
	book.close :savechanges => false
end

open('c:\temp\集計結果.csv','w') do |f|
f.puts %w(診療科 月 区分 社国 主治医 行為 理由 内容).join(",").encode('cp932')
data.each do |d|
	f.puts [d[:shinryouka],d[:month],d[:category],d[:kokuho],d[:shujii],d[:koui],d[:riyu],d[:kensa]].join(",").encode('cp932')
end
end

xl.quit
msgbox("出力完了しました。","完了")
system('explorer c:\temp\集計結果.csv')