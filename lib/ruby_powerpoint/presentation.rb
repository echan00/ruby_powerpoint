require 'zip'
require 'zip/filesystem'
require 'nokogiri'

module RubyPowerpoint

  class RubyPowerpoint::Presentation

    attr_reader :files

    def initialize path
      raise 'Not a valid file format.' unless (['.pptx'].include? File.extname(path).downcase)
      @files = Zip::File.open path
      @replace = {}
    end

    def slides
      slides = Array.new
      @files.each do |f|
        if f.name.include? 'ppt/slides/slide'
          slides.push RubyPowerpoint::Slide.new(self, f.name)
        end
      end
      slides.sort{|a,b| a.slide_num <=> b.slide_num}
    end

    def close
      @files.close
    end
    
    def save_and_return(slides)
      self.slides.each_with_index do |slide, index|
        @replace["slide"+slide.slide_num.to_s+".xml"] = slides[index].ret_slide_xml
      end
      puts @replace
      stringio = Zip::OutputStream.write_buffer do |out|
        @files.each do |entry|
          out.put_next_entry(entry.name)

          if @replace[entry.name]
            out.write(@replace[entry.name])
          else
            out.write(@files.read(entry.name))
          end
        end
      end
      @files.close
      return stringio
    end    
    
  end
end
