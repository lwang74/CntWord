Encoding.default_internal = "utf-8"
# Encoding.default_external = "gbk"

require 'yaml'
require_relative 'word'
require 'fileutils'

class Summary
	attr :all
	def initialize cfg
		@cfg = cfg
		doc = DocTable.new
		@all = []
		Dir["**/*.doc"].each{|one|
			if File.file?(one)
				if one =~ /\/([^\/~]+\.docx?)$/i
					file_name = $1
					puts file_name
					STDOUT.flush
					file_path="#{Dir.pwd}\\#{one}".gsub(/\//, "\\")
					result = doc.get_content(file_path)
					if result
						result['file_name'] = file_name
						result['file_path'] = file_path
						@all << result
					else
						puts "'#{one}' 有错误"
					end
				end
			end
		}
#~ p @all
		doc.close
	end

	def out file
		puts "*** Output ***"
		File.open(file, 'w'){|fout|
			if @all.size>0
				fout.puts @all[0].map{|k, v| k}.join(',') 
				@all.each{|one|
					fout.puts one.values.join(',')
				}
			end
		}
		rescue Exception => detail
			puts "先关闭#{file}'!"
			puts detail
	end
	
	def rename
		puts "*** ÎÄ¼þ¸ÄÃû ***"
		@all.each{|kemu, value|
			value.each{|one|
				puts one['ÐÕÃû']
				if one['file_path']=~/(.+\\).+(\..+)$/i
					path=$1
					ext=$2
					#~ puts path
					src=one['file_path']
					dst="#{path}#{one['ÐÕÃû']}#{ext}"
					FileUtils.mv src, dst if src!=dst 
				else
					puts "Â·¾¶ÃûÓÐ´í:'#{one['file_path']}'!"
				end
			}
		}
	end
	
	protected
	def date_chk dt
		if dt =~ /^(\d{2,4})([\.,\-]|£®|¡¢|Äê)(\d{1,2})([\.,\-]|£®|¡¢|ÔÂ)?(\d{1,2})?$/
			"#{$1}Äê#{$3}ÔÂ"
		else
			"=>#{dt}"
		end
	end	
	
	def work_age dt
		if dt =~ /^(\d+)(Äê)?$/
			"#{$1}Äê"
		else
			"=>#{dt}"
		end
	end	
end

def main
	doc = DocTable.new

	detect = YAML::load_file('detect.yml')
	sum = Summary.new(detect)
	sum.out 'Total.csv'
#~ sum.rename


	# Dir["**/*.doc"].each{|one|
	# 	if File.file?(one)
	# 		if one =~ /\/([^\/~]+\.docx?)$/i
	# 			file_name = $1
	# 			puts file_name
	# 			STDOUT.flush
	# 			file_path="#{Dir.pwd}\\#{one}".gsub(/\//, "\\")
	# 			result = doc.get_content(file_path)
	# 			if result
	# 				result['file_name'] = file_name
	# 				result['file_path'] = file_path
	# 			else
	# 				puts "'#{one}' 有错误"
	# 			end
	# 			p result
	# 		end
	# 	end
	# }

end

if ARGV.size==0 #读结果并写入CSV文件
	main
elsif ARGV.size==1	#从模板取detect
	#read template
	doc = DocTable.new
	doc.count_temp("#{Dir.pwd}\\#{ARGV[0]}".gsub(/\//, "\\"))
	doc.close
	puts "Write 'detect.yml' OK!"
else
	puts "Usage：CntWord.exe [Temp.doc]!"
end














