require 'slack-ruby-client'
require 'rubygems'
require 'write_xlsx'

# Class that helps to search user and create and update doc
class UserInfo
  def initialize_vars
    @row = 0
    @col = -1
    puts 'Userinfo search called'
  end

  def upload_file(client)
    client.files_upload(
      channels: '#general',
      as_user: true,
      file: Faraday::UploadIO.new('./userinfo.xlsx', 'userinfo.xlsx'),
      title: 'User Info',
      filename: 'userinfo.xlsx',
      initial_comment: 'User Info in Excel'
    )
  end

  def display_error(client, search)
    client.chat_postMessage(channel: '#general',
                            text: "Sorry, UID [#{search}] is invalid!",
                            as_user: true)
  end

  def display_intro(client)
    client.chat_postMessage(channel: '#general',
                            text: 'User Information', as_user: true)
  end

  def write_info(worksheet, key, value)
    worksheet.write(@row, @col += 1, key)
    worksheet.write(@row, @col += 1, value)
    @row += 1
    @col = -1
  end

  def search_user(client, worksheet, search)
    userinfo = client.users_info(user: search)
    userinfo['user'].each do |key, value|
      if value.class != Slack::Messages::Message
        write_info(worksheet, key, value)
      else
        userinfo['user'][key].each do |key2, value2|
          write_info(worksheet, key2, value2)
        end
      end
    end
  end

  def search_and_handle(client, worksheet, search, workbook)
    search_user(client, worksheet, search)
    workbook.close
    upload_file(client)
  rescue StandardError
    display_error(client, search)
    workbook.close
  end

  def get_user_info(client, search)
    puts "Search userinfo called for #{search}"
    initialize_vars
    display_intro(client)
    workbook = WriteXLSX.new('userinfo.xlsx')
    worksheet = workbook.add_worksheet
    search_and_handle(client, worksheet, search, workbook)
  end
end
