require 'slack-ruby-client'
require 'rubygems'
require 'write_xlsx'

# Class that helps to search channel info and upload itc
class ChannelInfo
  def initialize_vars
    @row = 0
    @col = -1
    puts 'Channelinfo search called'
  end

  def upload_file(client)
    client.files_upload(
      channels: '#general',
      as_user: true,
      file: Faraday::UploadIO.new('./channelinfo.xlsx', 'channelinfo.xlsx'),
      title: 'Channel Info',
      filename: 'channelinfo.xlsx',
      initial_comment: 'Channel Info in Excel'
    )
  end

  def display_error(client, search)
    client.chat_postMessage(channel: '#general',
                            text: "Sorry, channel [##{search}] is invalid!",
                            as_user: true)
  end

  def display_intro(client)
    client.chat_postMessage(channel: '#general',
                            text: 'Channel Information', as_user: true)
  end

  def write_info(worksheet, key, value)
    worksheet.write(@row, @col += 1, key)
    worksheet.write(@row, @col += 1, value)
    @row += 1
    @col = -1
  end

  def search_channel(client, worksheet, search)
    channelinfo = client.channels_info(channel: "##{search}")
    channelinfo['channel'].each do |key, value|
      if value.class != Slack::Messages::Message
        write_info(worksheet, key, value)
      else
        channelinfo['channel'][key].each do |key2, value2|
          write_info(worksheet, key2, value2)
        end
      end
    end
  end

  def search_and_handle(client, worksheet, search, workbook)
    search_channel(client, worksheet, search)
    workbook.close
    upload_file(client)
  rescue StandardError
    display_error(client, search)
    workbook.close
  end

  def get_channel_info(client, search)
    puts "Search channelinfo called for #{search}"
    initialize_vars
    display_intro(client)
    workbook = WriteXLSX.new('channelinfo.xlsx')
    worksheet = workbook.add_worksheet
    search_and_handle(client, worksheet, search, workbook)
  end
end
