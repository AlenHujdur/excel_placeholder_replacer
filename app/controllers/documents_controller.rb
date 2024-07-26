class DocumentsController < ApplicationController
  require 'rubyXL'
  require 'rubyXL/convenience_methods/cell'
  require 'rubyXL/convenience_methods/color'
  require 'rubyXL/convenience_methods/font'
  require 'rubyXL/convenience_methods/workbook'
  require 'rubyXL/convenience_methods/worksheet'

  def index
    @documents = Document.all
  end
  def new
    @document = Document.new
  end

  def create
    @document = Document.new(document_params)
    if @document.save
      DocProcessor.new(@document).process
      redirect_to @document
    else
      render :new
    end
  end

  def show
    @document = Document.find(params[:id])
    @rendered_document = DocProcessor.new(@document).render_document

    if @document.processed_document_path.present? && (@document.processed_document_path.include?('.xlsx') || @document.processed_document_path.include?('.xls'))
      path = Rails.root.join('public', @document.processed_document_path.gsub(%r{\A/+}, ''))
      puts "Opening processed document at path: #{path}"

      if File.exist?(path)
        puts "File exists at path: #{path}"
        @html_table = excel_to_html_with_styles(path)
      else
        puts "File does not exist at path: #{path}"
        @html_table = "<p>File does not exist at path: #{path}</p>"
      end
    else
      @html_table = "<p>Invalid or missing document path.</p>"
    end
  end

  def excel_to_html_with_styles(file_path)
    workbook = RubyXL::Parser.parse(file_path)
    html = ""

    workbook.worksheets.each do |worksheet|
      html << "<h2>#{worksheet.sheet_name}</h2>"
      html << "<table border='1'>"
      worksheet.each do |row|
        html << "<tr>"
        row && row.cells.each do |cell|
          html << "<td#{style_to_html(cell)}>#{cell && cell.value}</td>"
        end
        html << "</tr>"
      end
      html << "</table>"
    end

    html
  end

  def style_to_html(cell)
    return "" unless cell

    styles = []
    font_name = cell.font_name
    font_size = cell.font_size
    font_color = cell.font_color
    fill_color = cell.fill_color
    bold = cell.is_bolded
    italic = cell.is_italicized
    h_align = cell.horizontal_alignment
    v_align = cell.vertical_alignment

    styles << "font-family:#{font_name};" if font_name
    styles << "font-size:#{font_size}px;" if font_size
    styles << "color:##{font_color};" if font_color
    styles << "background-color:##{fill_color};" if fill_color
    styles << "font-weight:bold;" if bold
    styles << "font-style:italic;" if italic
    styles << "text-align:#{h_align};" if h_align
    styles << "vertical-align:#{v_align};" if v_align

    styles.any? ? " style=\"#{styles.join(' ')}\"" : ""
  end

  private

  def document_params
    params.require(:document).permit(:document)
  end
end
