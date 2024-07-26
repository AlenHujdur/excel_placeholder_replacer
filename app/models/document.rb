class Document < ApplicationRecord
  mount_uploader :document, DocumentUploader
end
