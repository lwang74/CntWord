require 'win32ole'
require 'yaml'

class DocTable
	def initialize points
		@points = points
		@word = WIN32OLE.new('Word.Application')
	end
	
	def get_content doc_file
		@doc = @word.Documents.open(doc_file, 'ReadOnly' => true)
		@doc.tables.each{|tb|
			#~ (1..tb.rows.count).each{|ri|
				#~ (1..tb.columns.count).each{|ci|
					#~ puts tb.cell(ri, ci).range.text
				#~ }
			#~ }
			tb.rows.each(){|r|
			}
			
			#~ }
			#~ .item{|row|
				#~ p row
			#~ }
			#~ .each{|row|
				#~ p row
			#~ }
		}
		result = {}
		#~ @points['table'].each{|key, val|
			#~ p key
			#~ p val
			#~ begin
			#~ tb = @doc.tables(val[0])
			#~ result[key] = trim(tb.cell(val[1], val[2]).range.text)
			#~ if val[3]
				#~ re = Regexp.new(val[3])
				#~ re =~ result[key]
				#~ result[key] = $1
			#~ end
			#~ ##~ puts result[key]
			#~ ##~ p result[key]
			#~ rescue Exception => detail
			#~ puts ">>> #{detail} <<<"
			#~ puts "#{key}->#{val}"
			#~ end
		#~ }
		
		#~ @points['para'].each{|key, val|
			#~ para_cnt = val[0].to_i
			#~ result[key] = ''
			#~ cnt =100
			#~ while cnt>0
				#~ startPara = trim(@doc.Paragraphs(para_cnt).range.text)
			##puts startPara
			#~ if val[1] and Regexp.new(val[1]) =~ startPara
			  #~ result[key] = $1
			  #~ break
			#~ end
			#~ para_cnt+=1
			#~ cnt-=1
		      #~ end
			##puts result[key]
		#~ }		
		@doc.close
		
		@points['const'].each{|key, val|
			result[key] = val
		}		
		#~ result
		
		#~ rescue Exception => detail
		#~ puts ">>>#{detail}"
		#~ nil
	end
	
	def close
		@word.quit
	end
	
	protected
	def trim val
		val.gsub(/[\s\r\a\?]/, '').gsub(/[,]/, '，')
	end
end

def dump result
	result.each{|k, v|
		puts "#{k}: #{v}"
	}
end

if __FILE__==$0
	all = YAML::load_file('config.yml')
	doc = DocTable.new(all['worddoc'])
	
	puts result = doc.get_content("E:\\Lyx\\CntWord\\temp.doc")
	
	doc.close
end


