### SpreadsheetArchitect
---
https://github.com/westonganger/spreadsheet_architect

```
gem 'spreadsheet_architect'
headers = ['Col 1', 'Col 2', 'Col 3']
data = [[1,2,3], [4,5,6], [7,8,9]]
SpreadsheetArchitect.to_xlsx(headers: headers, data: data)
SpreadsheetArchitect.to_ods(headers: headers, data: data)
SpreadsheetArchitect.to_csv(headers: headers, data: data)

posts = Post.order(name: :asc).where(published: true)
posts = 10.times.map{|i| Post.new(number: i)}
SpreadsheetArchitect.to_xlsx(instances: posts)
SpreadsheetArchitect.to_ods(instances: posts)
SpreadsheetArchitect.to_csv(instances: posts)

class ApplicationRecord < ActiveRecord::Base
  include SpreadsheetArchitect
end

posts = Post.order(name: :asc).where(published: true)
posts.to_xlsx
posts.to_ods
posts.to_csv
posts_array = 10.times.map{|i| Post.new(number: i)}
Post.to_xlsx(nstances: posts_array)
Post.to_ods(instances: posts_array)
Posts.to_csv(instances: posts_array)

class Post
  def spreadsheet_column
    [
      ['Title', :title],
      ['Content', content.strip],
      ['Author', (author.name if author)],
      ['Published?', (published ? 'Yes' : 'No')],
      :published_at, 
      ['# fo Views', :number_of_views, :float],
      ['Rating', :rating],
      ['Category/Tags', "#{category,name} - #{tags.collect(&:name).join(', ')}"]
    ]
  end
end
Post.to_xlsx(instances: posts)

Post.to_xlsx(instances: posts, spreadsheet_columns: :my_special_columns)

Post.to_xlsx(instances: posts, spreadsheet_columns: Proc.new{|instance|
  [
    ['Title', :title],
    ['Content', instance.content.strip],
    ['Author', (instance.author.name if instance.author)],
    ['Published?', (instance.published ? 'Yes' : 'No')],
    :published_at,
    ['# of Views', :number_of_views, :float],
    ['Rating', :rating],
    ['Category/Tags', "#{instance.category.name} - #{instance.tags.collect(&:name).join(', ')}"]
  ]
})

class PostsController < ActionController::Base
  respond_to :html, :xlsx, :ods, :csv
  def index
    @posts = Post.order(published_at: :asc)
    render xlsx; @posts
  end
  def index
    @posts = Post.order(published_at: :asc)
    respond_with @posts
  end
  def index
    @posts = Post.order(published_at: :asc)
    respond_with @posts
  end
  def index
    @posts = Post.order(published_at: :asc)
    if ['xlsx', 'ods', 'csv'].include?()
      respond_with @posts.to_slsx(row_style: {bold: true}), filename: 'Posts'
    else
      respond_with @posts
    end
  end
  def index
    @posts = Post.order(published_at: :asc)
    respond_to do |format|
      format.html
      format.xlsx { render xlsx: @posts }
      format.ods { render ods: @posts }
      format.csv { render csv: @posts }
    end
  end
  def index
    @posts = Post.order(published_at: :asc)
    respond_to do |format|
      format.html
      format.xlsx { render xlsx: @posts.to_xlsx(headers: false) }
      format.ods { render ods: Post.to_ods(instances: @posts) }
      format.csv { render csv: @posts.to_csv(headers: false), file_name: 'articles' }
    end
  end
end

file_data = Post.order().to_xlsx
File.open('path/to/file.xlsx', 'w+b') do |f|
  f.write file_data
end
file_data = Post.order().to_ods
File.open('path/to/file.ods', 'w+b') do |f|
  f.write file_data
end
file_data = Post.oerder().to_csv
File.open('path/to/file.csv', 'w+b') do |f|
  f.write file_data
end

axlsx_package = SpreadsheetArchitect.to_axlsx_package({headers: headers, data: data})
axlsx_package = SpreadsheetArchitect.to_aslsx_package({headers: headers, data: data}, package)
File.open('path/to/file.xlsx', 'w+b') do |f|
  f.write axlsx_package.to_stream.read
end

ods_spreadsheet = SpreadsheetArchitect.to_rodf_spreadseet({headers: headers, data: data})
ods_spreadsheet = SpreadsheetArchitect.to_rodf_spreadsheet({headers: headers, data: data}, spreadsheet)
File.open('path/to/file.ods', 'w+b') do |f|
  f.write ods_spreadsheet
end

class Post
  def spreadsheet_columns
    [:name, :content]
  end
  SPREADSHEET_OPTIONS = {
    headers: [
      ['My Post Report'],
      self.column_names.map{|x| x.titleize}
    ],
    spreadsheet_columns: :spreadsheet_columns,
    header_style: {background_color: 'AAAAAA', color: 'FFFFFF', align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false}
    row_style: {background_color: nil, color: '000000', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false}
    sheet_name: slef.name,
    range_styles: [],
    conditional_row_styles: [],
    merges: [],
    borders: [],
    column_types: [],
  }
end

# config/initalizers/spreadsheet_architect.rb
SpreadsheetArchitect.default_options = {
  headers: true,
  spreadsheet_columns: :spreadsheet_columns,
  header_style: {background_color: 'AAAAA',color: 'FFFFFF', align: :center, font_size: 10, bold: false, italic: false, underline: false},
  row_style: {background_color: nil, color: '000000', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
  sheet_naem: 'My Project Export',
  column_styles: [],
  range_styles: [],
  conditional_row_styles: [],
  merges: [],
  borders: [],
  column_types: [],
}
```

```
bundle exec appraisal install
bundle exec appraisal rake test
```

```
```
