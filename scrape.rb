require 'mechanize'
require 'uri'
require 'axlsx'

class Vcard
  attr_accessor :name, :description, :profession, :address, :phone, :email, :website

  def description=(description)
    @description = description.gsub(/[\n\t]/, '').strip
  end

  def address=(address)
    @address = address.to_s.gsub(/[\n\t]/, '').strip
  end

  def phone=(phone)
    @phone = phone.join ', '
  end

  def website=(website)
    @website = website.gsub(/[\n\t]/, '').strip
  end
end

p = Axlsx::Package.new
wb = p.workbook
wb_header = %w( Név Leírás Foglalkozás Cím Telefon Email Weboldal )

agent = Mechanize.new
_node_to_parse = '.listItemDetails'

counties = [
  'Bács-Kiskun megye',
  'Baranya megye',
  'Békés megye',
  'Borsod-Abaúj-Zemplén megye',
  'Csongrád megye',
  'Fejér megye',
  'Győr-Moson-Sopron megye',
  'Hajdú-Bihar megye',
  'Heves megye',
  'Jász-Nagykun-Szolnok megye',
  'Komárom-Esztergom megye',
  'Nógrád megye',
  'Pest megye',
  'Somogy megye',
  'Szabolcs-Szatmár-Bereg megye',
  'Tolna megye',
  'Vas megye',
  'Veszprém megye',
  'Zala megye'
]

puts 'starting to scrape county info'

counties.each do |county|
  puts "scraping #{county}"

  vcards = []
  _where = URI.encode county
  base_url = "http://aranyoldalak.hu/kereses.jspv?what=magánóvoda&where=#{_where}"
  _page = 0

  loop do
    puts "on page #{_page + 1}"

    sleep(2) # act as a friendly user, not as an agressive crawler bot
    url = base_url
    url += "&page=#{_page}" if _page > 0
    puts "url: #{url}"
    page = agent.get url

    page.search(_node_to_parse).each do |item|
      vcard = Vcard.new

      vcard.name = item.css('.org a').inner_html
      vcard.description = item.css('p.description').inner_html
      vcard.profession = item.css('ul.profession a').inner_html
      vcard.address = item.css('.vcard p.address').children()[2]

      email_node = item.css('.vcard .emailValue a.fetchByClick')
      vcard.email = email_node.attribute('data-finalvalue').value unless email_node.empty?

      phone_node = item.css('.vcard .phoneValue span.fetchByClick')
      vcard.phone = phone_node.map {|n| n.attribute('data-finalvalue').value unless n.nil? } unless phone_node.empty?

      website_node = item.css('.vcard .webLinkValue a')
      vcard.website = website_node.inner_html unless website_node.empty?

      vcards << vcard
    end

    _page += 1

    more_page = page.search('//*[@id="topPagerNextBtn"]/span[1]').to_s
    break if more_page.empty?
  end

  # save vcards
  puts 'creating worksheet'

  wb.add_worksheet(:name => county) do |sheet|
    sheet.add_row wb_header
    vcards.each do |vcard|
      sheet.add_row [vcard.name, vcard.description, vcard.profession, vcard.address, vcard.phone, vcard.email, vcard.website]
    end
  end
end

print 'dumping vcards to workbook...'
p.serialize('db.xlsx')
puts 'done'
