require 'yaml'
require './word'
require 'fileutils'

class Summary
	attr :all
	def initialize cfg
		@cfg = cfg
		doc = DocTable.new(@cfg['worddoc'])
		@all = {}
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
						ord = result[@cfg['output']['order']]
						ord='' if !ord
						@all[ord] ||= []
						@all[ord] << result
					else
						puts "'#{one}' 格式错误！！！请检查"
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
			fout.puts @cfg['output']['cols'].join(',')
			#~ @all.each{|k, v|
				#~ puts "key=>#{k};value=>#{v}"
			#~ }
			@all.sort.each{|cat_k, cat_v|
				puts cat_k
				cat_v.each{|one|
					line = []
					@cfg['output']['cols'].each{|col|
						line << one[col]
					}
					fout.puts line.join(',')
				}
			}
		}
		rescue Exception => detail
			puts "请先关闭'#{file}'!"
			puts detail
	end
	
	def rename
		puts "*** 文件改名 ***"
		@all.each{|kemu, value|
			value.each{|one|
				puts one['姓名']
				if one['file_path']=~/(.+\\).+(\..+)$/i
					path=$1
					ext=$2
					#~ puts path
					src=one['file_path']
					dst="#{path}#{one['姓名']}#{ext}"
					FileUtils.mv src, dst if src!=dst 
				else
					puts "路径名有错:'#{one['file_path']}'!"
				end
			}
		}
	end
	
	protected
	def date_chk dt
		if dt =~ /^(\d{2,4})([\.,\-]|．|、|年)(\d{1,2})([\.,\-]|．|、|月)?(\d{1,2})?$/
			"#{$1}年#{$3}月"
		else
			"=>#{dt}"
		end
	end	
	
	def work_age dt
		if dt =~ /^(\d+)(年)?$/
			"#{$1}年"
		else
			"=>#{dt}"
		end
	end	
end

#~ cfg = YAML::load_file('config.yml')
#~ sum = Summary.new(cfg)
#~ sum.out 'Total.csv'
#~ sum.rename

Dir["**/*.doc"].each{|one|
	if File.file?(one)
		if one =~ /\/([^\/~]+\.docx?)$/i
			file_name = $1
			puts file_name
			STDOUT.flush
			file_path="#{Dir.pwd}\\#{one}".gsub(/\//, "\\")
			result = doc.get_content(file_path)
			#~ if result
				#~ result['file_name'] = file_name
				#~ result['file_path'] = file_path
				#~ ord = result[@cfg['output']['order']]
				#~ ord='' if !ord
				#~ @all[ord] ||= []
				#~ @all[ord] << result
			#~ else
				#~ puts "'#{one}' 格式错误！！！请检查"
			#~ end
		end
	end
}















