class PlaceholdersController < ApplicationController
  def index
    @placeholders = Placeholder.all
  end

  def new
    @placeholder = Placeholder.new
  end

  def create
    @placeholder = Placeholder.new(placeholder_params)
    if @placeholder.save
      redirect_to placeholders_path, notice: 'Placeholder created successfully'
    else
      render :new
    end
  end

  def edit
    @placeholder = Placeholder.find(params[:id])
  end

  def update
    @placeholder = Placeholder.find(params[:id])
    if @placeholder.update(placeholder_params)
      redirect_to placeholders_path, notice: 'Placeholder updated successfully'
    else
      render :edit
    end
  end

  def destroy
    @placeholder = Placeholder.find(params[:id])
    @placeholder.destroy
    redirect_to placeholders_path, notice: 'Placeholder deleted successfully'
  end

  private

  def placeholder_params
    params.require(:placeholder).permit(:key, :value)
  end
end
