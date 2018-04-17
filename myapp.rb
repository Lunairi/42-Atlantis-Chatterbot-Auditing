require 'slack-ruby-client'
require 'rubygems'
require 'write_xlsx'
require_relative 'userinfo.rb'
require_relative 'channelinfo.rb'

def respond_userinfo(client, data, wclient, get_user, input)
  search = input.sub('<@UA7PZSA4V> userinfo ', '')
  get_user.get_user_info(wclient, search)
  client.message channel: data.channel, text: 'User info request completed.'
end

def respond_chaninfo(client, data, wclient, get_channel, input)
  search = input.sub('<@UA7PZSA4V> channelinfo ', '')
  get_channel.get_channel_info(wclient, search)
  client.message channel: data.channel, text: 'Channel info request completed.'
end

def respond_help(client, data)
  client.message channel: data.channel, text:
  "Here are the available options <@#{data.user}>"
  client.message channel: data.channel, text:
  '[Get User Information] @chatterbot userinfo <UID>'
  client.message channel: data.channel, text:
  '[Get Channel Information] @chatterbot channelinfo <channel>'
end

def check_search_request(data, client, wclient, name)
  if data.text.include? "#{name} userinfo "
    respond_userinfo(client, data, wclient, UserInfo.new, data.text)
  elsif data.text.include? "#{name} channelinfo "
    respond_chaninfo(client, data, wclient, ChannelInfo.new, data.text)
  end
end

def respond_to_message(data, client, wclient, name)
  if data.text == "#{name} hi" || data.text == "#{name} hello"
    client.message channel: data.channel, text: "Hi there <@#{data.user}>!"
  elsif data.text == "#{name} help"
    respond_help(client, data)
  else
    check_search_request(data, client, wclient, name)
  end
end

def client_reaction(client)
  client.on :hello do
    puts "Chatterbot started on '#{client.team.name}'."
  end
  client.on :closed do |_data|
    puts 'Chatterbot has disconnected successfully!'
  end
end

def main
  Slack.configure do |config|
    config.token = 'xoxb-347815894165-HGTjjdCIwEybnbjMXTXPxf2H'
  end
  client = Slack::RealTime::Client.new
  wclient = Slack::Web::Client.new
  client_reaction(client)
  client.on :message do |data|
    respond_to_message(data, client, wclient, '<@UA7PZSA4V>')
  end
  client.start!
end

main
