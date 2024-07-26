class DocumentsController < ApplicationController
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
      # Correctly construct the full path to the processed document
      processed_document_path = @document.processed_document_path.gsub(%r{\A/+}, '') # Remove leading slash if present
      full_path = Rails.root.join('public', processed_document_path)

      puts "Opening processed document at path: #{full_path}"

      if File.exist?(full_path)
        puts "File exists at path: #{full_path}"
        @html_table = excel_to_html(full_path)
      else
        puts "File does not exist at path: #{full_path}"
        @html_table = "<p>File does not exist at path: #{full_path}</p>"
      end
    else
      @html_table = "<p>Invalid or missing document path.</p>"
    end
  end

  def excel_to_html(file_path)
    workbook = RubyXL::Parser.parse(file_path)
    html = ""

    workbook.worksheets.each do |worksheet|
      html << "<h2>#{worksheet.sheet_name}</h2>"
      html << "<table border='1'>"
      worksheet.each do |row|
        html << "<tr>"
        row && row.cells.each do |cell|
          html << "<td>#{cell && cell.value}</td>"
        end
        html << "</tr>"
      end
      html << "</table>"
    end

    html
  end

  private

  def document_params
    params.require(:document).permit(:document)
  end
end
