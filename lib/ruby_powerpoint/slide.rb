require 'zip/filesystem'
require 'nokogiri'

module RubyPowerpoint
  class RubyPowerpoint::Slide

    attr_reader :presentation,
                :slide_number,
                :slide_number,
                :slide_file_name

    def initialize presentation, slide_xml_path
      @presentation = presentation
      @slide_xml_path = slide_xml_path
      @slide_number = extract_slide_number_from_path slide_xml_path
      @slide_notes_xml_path = "ppt/notesSlides/notesSlide#{@slide_number}.xml"
      @slide_file_name = extract_slide_file_name_from_path slide_xml_path

      parse_slide
      parse_slide_notes
      parse_relation
    end

    def parse_slide
      slide_doc = @presentation.files.file.open @slide_xml_path
      @slide_xml = Nokogiri::XML::Document.parse slide_doc
    end

    def parse_slide_notes
      slide_notes_doc = @presentation.files.file.open @slide_notes_xml_path rescue nil

      if slide_notes_doc
        @slide_notes_xml = Nokogiri::XML::Document.parse(slide_notes_doc)
      end       
    end

    def parse_relation
      @relation_xml_path = "ppt/slides/_rels/#{@slide_file_name}.rels"
      if @presentation.files.file.exist? @relation_xml_path
        relation_doc = @presentation.files.file.open @relation_xml_path
        @relation_xml = Nokogiri::XML::Document.parse relation_doc
      end
    end

    def content
      content_elements @slide_xml
    end

    def notes_content      
      if @slide_notes_xml
        content_elements @slide_notes_xml
      else
        nil
      end     
    end

    def title
      title_elements = title_elements(@slide_xml)
      title_elements.join(" ") if title_elements.length > 0
    end

    def change_title(new_title, old_title, result)
      if(title == old_title)

        # Find the title
        temp = nil
        @slide_xml.xpath('//p:sp').each do |node|
          if(element_is_title(node))
            node.xpath('//a:t').each do |attempt|
              if(attempt.content == old_title)
                puts attempt.content
                attempt.content = new_title
              end
            end
          end
        end

        # # Write to file
        # @presentation.files.get_output_stream(@slide_xml_path) { |f| f.puts @slide_xml } 
        # outputstream = @presentation.files.get_output_stream(@slide_xml_path)
        # outputstream.write @slide_xml
        # outputstream.close
        # puts 'this is important operation'

        # Zip::ZipFile.open("spec/fixtures/rime.pptx", "wb") {
        #   |f| 
        #   os = f.get_output_stream(@slide_xml_path)
        #   os.write @slide_xml.to_s
        #   os.close
        #   f.commit
        # }

        # @presentation.files.get_output_stream(@slide_xml_path) {|f| f.write(@slide_xml.to_s)}  

        # buffer = Zip::ZipOutputStream.write_buffer do |out|
        #   out.put_next_entry(@slide_xml_path)
        #   out.write @slide_xml
        # end

        # @presentation.files.get_output_stream(@slide_xml_path) {|f| f.write(buffer.string) }

        # use dir.tmpdir

        # Rubyzip does not create a valid zip file in whatever way this is attempted
        # Alternative in commandline
        name = @presentation.files.name
        if(name.include? '/')
          folder = name[0..name.rindex('/')]
          result = folder + result
        end
      
        xmlFiles = 'docProps ppt _rels [Content_Types].xml'
        
        # unzip the pptx
        `unzip #{name}`

        # overwrite the necessary file
        File.open(@slide_xml_path, 'w+') { |f| f.write(@slide_xml.to_s) }

        # zip the pptx
        `zip #{result} -r #{xmlFiles}`

        # remove the folders
        `rm -rf #{xmlFiles}`
        return
      end
    end

    def update_texts(old_texts, new_texts)
      # Find the title
      temp = nil
      old_texts.each_with_index do |old_text, idx|
        @slide_xml.xpath('//p:sp').each do |node|
          node.xpath('//a:t').each do |attempt|
            if(attempt.content == old_text)
              attempt.content = new_texts[idx]
            end
          end
        end
      end

      result = @slide_file_name[0...-4]+"-converted.pptx" 
      # Rubyzip does not create a valid zip file in whatever way this is attempted
      # Alternative in commandline
      name = @presentation.files.name
      if(name.include? '/')
        folder = name[0..name.rindex('/')]
        result = folder + result
      end

      xmlFiles = 'docProps ppt _rels [Content_Types].xml'

      # unzip the pptx
      `unzip #{name}`

      # overwrite the necessary file
      File.open(@slide_xml_path, 'w+') { |f| f.write(@slide_xml.to_s) }

      # zip the pptx
      `zip #{result} -r #{xmlFiles}`

      # remove the folders
      `rm -rf #{xmlFiles}`
      return
    end        
    
    def images
      image_elements(@relation_xml)
        .map.each do |node|
          @presentation.files.file.open(
            node['Target'].gsub('..', 'ppt'))
        end
    end

    def slide_num
      @slide_xml_path.match(/slide([0-9]*)\.xml$/)[1].to_i
    end

    def paragraphs
      paragraph_element @slide_xml
    end

    def paragraphs_xml
      @slide_xml.xpath('//a:p')
    end    

    def ret_slide_xml
      @slide_xml
    end      
    
    private

    def extract_slide_number_from_path path
      path.gsub('ppt/slides/slide', '').gsub('.xml', '').to_i
    end

    def extract_slide_file_name_from_path path
      path.gsub('ppt/slides/', '')
    end

    def title_elements(xml)
      shape_elements(xml).select{ |shape| element_is_title(shape) }
    end

    def content_elements(xml)
      xml.xpath('//a:t').collect{ |node| node.text }
    end

    def image_elements(xml)
      xml.css('Relationship').select{ |node| element_is_image(node) }
    end

    def shape_elements(xml)
      xml.xpath('//p:sp')
    end

    def paragraph_element(xml)
      xml.xpath('//a:p').collect{ |node| RubyPowerpoint::Paragraph.new(self, node) }
    end

    def element_is_title(shape)
      shape.xpath('.//p:nvSpPr/p:nvPr/p:ph').select{ |prop| prop['type'] == 'title' || prop['type'] == 'ctrTitle' }.length > 0
    end

    def element_is_image(node)
      node['Type'].include? 'image'
    end
  end
end
