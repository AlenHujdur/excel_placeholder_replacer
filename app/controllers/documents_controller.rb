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

  private

  def excel_to_html_with_styles(file_path)
    workbook = RubyXL::Parser.parse(file_path)
    html = ""

    workbook.worksheets.each do |worksheet|
      merges = prepare_merge_ranges(worksheet)
      html << "<h2>#{worksheet.sheet_name}</h2>"
      html << "<table border='1'>"
      worksheet.each_with_index do |row, row_idx|
        html << "<tr>"
        row && row.cells.each_with_index do |cell, col_idx|
          next if merges[[row_idx, col_idx]] == :skip
          html << "<td#{style_to_html(cell)}#{merge_html(merges, row_idx, col_idx)}>"
          html << "#{cell && cell.value}</td>"
        end
        html << "</tr>"
      end
      html << "</table>"
    end

    html
  end

  def merge_html(merges, row, col)
    merge = merges[[row, col]]
    return "" unless merge
    " colspan='#{merge[:colspan]}' rowspan='#{merge[:rowspan]}'"
  end

  def prepare_merge_ranges(worksheet)
    merges = {}
    worksheet.merged_cells.each do |range|
      puts "Merged range: #{range.inspect}"
      row_range = range.ref.row_range
      col_range = range.ref.col_range

      (row_range.first..row_range.last).each do |row|
        (col_range.first..col_range.last).each do |col|
          merges[[row, col]] = :skip
        end
      end

      merges[[row_range.first, col_range.first]] = { colspan: col_range.size, rowspan: row_range.size }
    end
    merges
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
    wrap_text = cell.change_text_wrap(true)

    # Check border lines
    top_border_line = cell.get_border(:top).to_s != ''
    left_border_line = cell.get_border(:left).to_s != ''
    right_border_line = cell.get_border(:right).to_s != ''
    bottom_border_line = cell.get_border(:bottom).to_s != ''

    top_border_style = cell.get_border(:top) == 'thin' ? '1px solid black' : ''
    left_border_style = cell.get_border(:left) == 'thin' ? '1px solid black' : ''
    right_border_style = cell.get_border(:right) == 'thin' ? '1px solid black' : ''
    bottom_border_style = cell.get_border(:bottom) == 'thin' ? '1px solid black' : ''

    styles << "border-top: #{top_border_style};" if top_border_line
    styles << "border-left: #{left_border_style};" if left_border_line
    styles << "border-right: #{right_border_style};" if right_border_line
    styles << "border-bottom: #{bottom_border_style};" if bottom_border_line

    styles << "font-family:#{font_name};" if font_name
    styles << "font-size:#{font_size}px;" if font_size
    styles << "color:##{font_color};" if font_color
    styles << "background-color:##{fill_color};" if fill_color
    styles << "font-weight:bold;" if bold
    styles << "font-style:italic;" if italic
    styles << "text-align:#{h_align};" if h_align
    styles << "vertical-align:#{v_align};" if v_align
    styles << "white-space: normal; word-wrap: break-word;" if wrap_text

    styles.any? ? " style=\"#{styles.join(' ')}\"" : ""
  end

  def document_params
    params.require(:document).permit(:document)
  end
end
