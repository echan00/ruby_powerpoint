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
      
      @slides = Array.new
      @diagrams = Array.new      
      @charts = Array.new      
      
      @files.each do |f|
        if f.name.include? 'ppt/slides/slide'
          @slides.push RubyPowerpoint::Slide.new(self, f.name, 0)
        end
        if f.name.include? 'ppt/diagrams/data'
          @diagrams.push RubyPowerpoint::Slide.new(self, f.name, 1)
        end
        if f.name.include? 'ppt/charts/chart'
          @charts.push RubyPowerpoint::Slide.new(self, f.name, 2)
        end       
      end
      @slides.sort{|a,b| a.slide_num <=> b.slide_num}
      @diagrams.sort{|a,b| a.slide_num <=> b.slide_num}
      @charts.sort{|a,b| a.slide_num <=> b.slide_num}      
    end

    def slides
      return @slides
    end

    def diagrams
      return @diagrams
    end

    def charts
      return @charts
    end    
    
    def close
      @files.close
    end
    
    def save_and_return(old_slides, old_diagrams, old_charts)      
      @slides.each_with_index do |slide, index|
        @replace["ppt/slides/slide"+slide.slide_num.to_s+".xml"] = old_slides[index].ret_slide_xml.serialize(:save_with => 0)
      end
      @diagrams.each_with_index do |diagram, index|
        @replace["ppt/diagrams/data"+diagram.slide_num.to_s+".xml"] = old_diagrams[index].ret_slide_xml.serialize(:save_with => 0)
      end
      @charts.each_with_index do |chart, index|
        @replace["ppt/charts/chart"+chart.slide_num.to_s+".xml"] = old_charts[index].ret_slide_xml.serialize(:save_with => 0)
      end      
      stringio = Zip::OutputStream.write_buffer do |out|
        @files.each do |entry|
          out.put_next_entry(entry.name)

          puts entry.name
          if @replace[entry.name]
            out.write(@replace[entry.name])
          elsif entry.directory?
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
