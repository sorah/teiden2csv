#-*- coding: utf-8 -*-
require 'rubygems'
require 'spreadsheet'

options = {:disable_update => false,
           :cache => "#{File.dirname(__FILE__)}/cache",
           :filename => "teiden.csv"
          }

while arg = ARGV.shift
  case arg
  when /--cache=(.+)/
    options[:cache] = $1.gsub(/\/$/,"")
  when "--disable-update", "-du"
    options[:disable_update] = true
  else
    options[:filename] = arg
  end
end

FILE_URLS = %w{
  http://www.tepco.co.jp/images/tochigi.xls
  http://www.tepco.co.jp/images/ibaraki.xls
  http://www.tepco.co.jp/images/gunma.xls
  http://www.tepco.co.jp/images/chiba.xls
  http://www.tepco.co.jp/images/kanagawa.xls
  http://www.tepco.co.jp/images/tokyo.xls
  http://www.tepco.co.jp/images/saitama.xls
  http://www.tepco.co.jp/images/yamanashi.xls
  http://www.tepco.co.jp/images/numazu.xls
}
FILES = %w{
  tochigi.xls
  ibaraki.xls
  gunma.xls
  chiba.xls
  kanagawa.xls
  tokyo.xls
  saitama.xls
  yamanashi.xls
  numazu.xls
}
FILES.map!{|x|x.prepend(options[:cache]+"/")}
ARROW = "\e[1m\e[34m=>\e[0m\e[1m\e[0m"

if File.exist?(options[:cache])
  abort "Cache directory is exist and not directory" unless FileTest.directory?(options[:cache])
else
  Dir.mkdir(options[:cache])
end

CURL = !(system("curl", err: File::NULL).nil?)
require 'open-uri' unless CURL

FILE_URLS.zip(FILES).each do |url,file|
  if File.exist?(file) && options[:disable_update]
    puts "#{ARROW} Skipping #{url}"
  else
    puts "#{ARROW} Downloading #{url} to #{file}"
    if CURL
      system "curl","-#","-o",file,url
    else
      open(file,"wb"){|io| io.write open(url).read }
    end
  end
end

array = FILES.map do |file|
  puts "#{ARROW} #{file}: converting"
  book = Spreadsheet.open(file,'rb')
  sheet = book.worksheet(0)
  sheet.each_with_index.map do |v,i|
    next if i == 0
    next if v[0].include?("なお、") || v[0].include?("都県")
    [v[0],v[1],v[2],v[3].to_i,v[4]].join(",")
  end.compact.join("\n")
end

puts "#{ARROW} Writing to #{File.expand_path(options[:filename])}"

open(options[:filename],"w") {|io| io.puts array.join("\n") }
