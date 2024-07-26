class CreateDocuments < ActiveRecord::Migration[7.1]
  def change
    create_table :documents do |t|
      t.string :document
      t.string :processed_document_path

      t.timestamps
    end
  end
end
