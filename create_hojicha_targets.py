import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Hojicha Export Targets"

headers = ["Country", "City", "Business Name", "Category", "Website", "Instagram", "Email", "Notes"]

# Styles
header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
header_fill = PatternFill(start_color="2D5016", end_color="2D5016", fill_type="solid")
alt_fill = PatternFill(start_color="F0F5EA", end_color="F0F5EA", fill_type="solid")
border = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC")
)
center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

# Write headers
for col, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_align
    cell.border = border

data = [
    # ===== SINGAPORE (30 stores) =====
    # Specialty Cafes
    ("Singapore", "Singapore", "Nylon Coffee Roasters", "Specialty Cafe", "https://nylon.coffee", "@nyloncoffee", "hello@nyloncoffee.com", "Award-winning micro-roastery; single-origin focus; hojicha menu potential"),
    ("Singapore", "Singapore", "Chye Seng Huat Hardware (PPP Coffee)", "Specialty Cafe", "https://pppcoffee.com", "@pppcoffee", "hello@pppcoffee.com", "Pioneer SG specialty cafe; roastery on-site; strong tea menu interest"),
    ("Singapore", "Singapore", "Sarnies", "Specialty Cafe", "https://sgcoffee.sarnies.com", "@sarniescoffee", "info@sarnies.com.sg", "CBD flagship; Australian-style café; curated tea & coffee menu"),
    ("Singapore", "Singapore", "The Coffee Academics", "Specialty Cafe", "https://www.theacademicsgroup.com", "@thecoffeeacademics", "info@theacademicsgroup.com", "48 stores across Asia; award-winning roasters; active Japanese tea interest"),
    ("Singapore", "Singapore", "Dutch Colony Coffee Co.", "Specialty Cafe", "https://www.dutchcolony.com", "@dutchcolonycoffee", "hello@dutchcolony.com", "Artisan roastery café; single-origin specialty; boutique tea offerings"),
    ("Singapore", "Singapore", "Common Man Coffee Roasters", "Specialty Cafe", "https://www.commonmancoffeeroasters.com", "@commonmancoffeeroasters", "info@cmcroasters.com", "Flagship roastery in Robertson Quay; premium tea & coffee menu"),
    ("Singapore", "Singapore", "Tiong Hoe Specialty Coffee", "Specialty Cafe", "https://tionghoe.com", "@tionghoespecialtycoffee", "hello@tionghoe.com", "Micro-roaster; single-origin focus; curated alternative beverages"),
    ("Singapore", "Singapore", "Homeground Coffee Roasters", "Specialty Cafe", "https://homegroundcoffeeroasters.com", "@homeground.coffee", "hello@homegroundcoffeeroasters.com", "Community-first roastery; specialty tea interest; Tiong Bahru location"),
    ("Singapore", "Singapore", "Strangers' Reunion", "Specialty Cafe", "https://www.strangersreunion.com", "@strangersreu", "enquiries@strangersreunion.com", "Tiong Bahru café icon; known for quality beverages; Japanese tea potential"),
    ("Singapore", "Singapore", "Proud Mary Coffee", "Specialty Cafe", "https://proudmarycoffee.com.au", "@proudmarycoffee", "singapore@proudmarycoffee.com", "Melbourne-origin specialty roaster; global expansion; premium tea interest"),
    # Tea Brands
    ("Singapore", "Singapore", "TWG Tea", "Tea Brand", "https://twgtea.com", "@twgtea", "retail@twgtea.com", "Founded in SG 2008; 1,000+ blends; 70+ boutiques globally; key partner prospect"),
    ("Singapore", "Singapore", "Gryphon Tea Company", "Tea Brand", "https://www.gryphontea.com", "@gryphontea", "info@gryphontea.com", "SG heritage brand since 1918; award-winning blends; sold in 6 Asian countries"),
    ("Singapore", "Singapore", "Pek Sin Choon", "Tea Brand", "https://www.peksinchoon.com", "@peksinchoon", "info@peksinchoon.com", "Oldest SG tea merchant since 1925; traditional & premium blends"),
    ("Singapore", "Singapore", "The 1872 Clipper Tea Co.", "Tea Brand", "https://www.1872clipperte.com.sg", "@1872clippertea", "hello@1872clipperte.com.sg", "SG family-run tea brand; curated premium blends; hojicha line potential"),
    ("Singapore", "Singapore", "Gryphon Tea Salon at Ion Orchard", "Tea Brand", "https://www.gryphontea.com", "@gryphontea", "salon@gryphontea.com", "Luxury tea salon concept; Orchard Road location; premium hojicha fit"),
    # Luxury Hotels
    ("Singapore", "Singapore", "Raffles Hotel Singapore", "Luxury Hotel", "https://www.raffles.com/singapore", "@raffleshottels", "singapore@raffles.com", "Iconic 5-star heritage hotel; 10 F&B outlets; high-end afternoon tea program"),
    ("Singapore", "Singapore", "Marina Bay Sands", "Luxury Hotel", "https://www.marinabaysands.com", "@marinabaysands", "hotel@marinabaysands.com", "Icon hotel & resort; diverse F&B; Louis Vuitton Chocolaterie on-site"),
    ("Singapore", "Singapore", "The Ritz-Carlton, Millenia Singapore", "Luxury Hotel", "https://www.ritzcarlton.com/singapore", "@ritzcarltonsingapore", "rc.sinrz.leads@ritzcarlton.com", "Luxury 5-star; Summer Pavilion restaurant; premium tea service"),
    ("Singapore", "Singapore", "Mandarin Oriental Singapore", "Luxury Hotel", "https://www.mandarinoriental.com/singapore", "@mo_singapore", "mosin-reservations@mohg.com", "5-star marina views; 5 restaurants; premium Oriental club lounge tea service"),
    ("Singapore", "Singapore", "Four Seasons Hotel Singapore", "Luxury Hotel", "https://www.fourseasons.com/singapore", "@fssngapore", "res.singapore@fourseasons.com", "Orchard Road location; upscale F&B; afternoon tea with TWG collaboration"),
    ("Singapore", "Singapore", "Shangri-La Singapore", "Luxury Hotel", "https://www.shangri-la.com/singapore", "@shangrila_sin", "sin.dining@shangri-la.com", "15-acre tropical retreat; premium dining; tea amenities for guests"),
    ("Singapore", "Singapore", "The Fullerton Hotel Singapore", "Luxury Hotel", "https://www.fullertonhotels.com/fullerton-hotel-singapore", "@fullertonhotelsingapore", "info@fullertonhotel.com.sg", "Heritage neo-classical landmark; The Courtyard afternoon tea experience"),
    ("Singapore", "Singapore", "Capella Singapore", "Luxury Hotel", "https://www.capellahotels.com/singapore", "@capellasingapore", "singapore@capellahotels.com", "Luxury Sentosa retreat; Auriga spa; premium culinary experiences"),
    ("Singapore", "Singapore", "The Singapore EDITION", "Luxury Hotel", "https://www.editionhotels.com/singapore", "@editionhotels", "singapore@editionhotels.com", "Opened 2024 Orchard Road; signature Black Tea fragrance; trendsetting F&B"),
    ("Singapore", "Singapore", "Conrad Singapore Orchard", "Luxury Hotel", "https://www.hilton.com/en/hotels/sinorci-conrad-singapore-orchard", "@conradsingapore", "conradsingaporedinings@conradhotels.com", "Opened 2024; 10 bars & restaurants incl. Michelin-starred Chinese restaurant"),
    # Luxury Restaurants
    ("Singapore", "Singapore", "Odette", "Luxury Restaurant", "https://www.odetterestaurant.com", "@odetterestaurant", "enquiry@odetterestaurant.com", "3 Michelin stars; Asia's 50 Best; French contemporary; premium tea pairing"),
    ("Singapore", "Singapore", "Les Amis", "Luxury Restaurant", "https://www.lesamis.com.sg", "@lesamis_sg", "reservations@lesamis.com.sg", "3 Michelin stars; classic French fine dining; curated beverage pairing menu"),
    ("Singapore", "Singapore", "Jaan by Kirk Westaway", "Luxury Restaurant", "https://www.jaan.com.sg", "@jaanbykirk", "reservations@jaan.com.sg", "1 Michelin star; British contemporary; 70th floor Swissôtel; tea pairings"),
    ("Singapore", "Singapore", "Burnt Ends", "Luxury Restaurant", "https://www.burntends.com.sg", "@burntends_sg", "reservations@burntends.com.sg", "1 Michelin star; Australia-style open-fire; adventurous beverage pairings"),
    ("Singapore", "Singapore", "Zén", "Luxury Restaurant", "https://restaurantzen.com", "@restaurant_zen", "reservations@restaurantzen.com", "3 Michelin stars; Swedish-Japanese inspired; premium tea & sake pairing"),
    # ===== HONG KONG (30 stores) =====
    # Specialty Cafes
    ("Hong Kong", "Hong Kong", "Blue Bottle Coffee Hong Kong", "Specialty Cafe", "https://bluebottlecoffee.com/cafes/hong-kong", "@bluebottlecoffee", "hongkong@bluebottlecoffee.com", "Global specialty coffee icon; ifc mall & Lyndhurst Terrace; premium tea interest"),
    ("Hong Kong", "Hong Kong", "Elephant Grounds", "Specialty Cafe", "https://www.elephantgrounds.com", "@elephantgrounds", "hello@elephantgrounds.com", "HK-born lifestyle brand since 2013; house-roasted beans; Japanese-influenced menu"),
    ("Hong Kong", "Hong Kong", "Knockbox Coffee Company", "Specialty Cafe", "https://www.knockboxcoffee.hk", "@knockboxcoffee", "info@knockboxcoffee.hk", "Award-winning local roaster; HK Barista Champion 2x; premium single-origin"),
    ("Hong Kong", "Hong Kong", "The Coffee Academics HK", "Specialty Cafe", "https://www.theacademicsgroup.com", "@thecoffeeacademics", "info@theacademicsgroup.com", "Pioneer HK specialty coffee since 2012; 17+ HK outlets; ISO-certified roastery"),
    ("Hong Kong", "Hong Kong", "The Cupping Room", "Specialty Cafe", "https://cuppingroom.hk", "@cuppingroom.hk", "hello@cuppingroom.hk", "Pioneer since 2011; World Barista Championship 2nd; harbor view café"),
    ("Hong Kong", "Hong Kong", "FINEPRINT", "Specialty Cafe", "https://fineprint.hk", "@fineprint.hk", "hello@fineprint.hk", "Coffee & sourdough concept; multiple HK locations incl. IFC & airport"),
    ("Hong Kong", "Hong Kong", "Holt's Café – Rosewood Hong Kong", "Specialty Cafe", "https://www.rosewoodhotels.com/en/hong-kong/dining/holts-cafe", "@rosewoodhongkong", "rhongkong@rosewoodhotels.com", "Stylish hotel café & tea house at Rosewood; premium Japanese tea opportunity"),
    ("Hong Kong", "Hong Kong", "Café Gray Deluxe – The Upper House", "Specialty Cafe", "https://www.upperhouse.com/en/dining/cafe-gray-deluxe", "@theupperhouse", "dining@upperhouse.com", "Upper House hotel café; panoramic city views; curated premium beverage menu"),
    ("Hong Kong", "Hong Kong", "The Tiffany Blue Box Café HK", "Specialty Cafe", "https://www.tiffany.com/blue-box-cafe", "@tiffanyandco", "hk.cafe@tiffany.com", "Luxury brand café; Tsim Sha Tsui; afternoon tea; premium Japanese tea fit"),
    ("Hong Kong", "Hong Kong", "Ralph's Coffee Hong Kong", "Specialty Cafe", "https://www.ralphlauren.com/ralphs-coffee", "@ralphscoffee", "hk@ralphscoffee.com", "Luxury fashion café; La Colombe roasted beans; upscale tea menu interest"),
    # Tea Brands
    ("Hong Kong", "Hong Kong", "Fortnum & Mason Hong Kong", "Tea Brand", "https://www.fortnumandmason.com", "@fortnumandmason", "hongkong@fortnumandmason.com", "300+ year British heritage; Royal Blend; K11 MUSEA Tsim Sha Tsui location"),
    ("Hong Kong", "Hong Kong", "Lock Cha Tea Shop", "Tea Brand", "https://www.lockcha.com", "@lockchateashop", "info@lockcha.com", "Renowned HK traditional tea house; specializes in quality Chinese teas"),
    ("Hong Kong", "Hong Kong", "Yú Tea House", "Tea Brand", "https://www.yuteahouse.com", "@yuteahouse", "hello@yuteahouse.com", "Curated Asian tea programs; in-hotel concepts at Four Seasons HK"),
    ("Hong Kong", "Hong Kong", "TWG Tea Hong Kong", "Tea Brand", "https://twgtea.com", "@twgtea", "hongkong@twgtea.com", "Luxury SG tea brand; HarbourCity & IFC locations; premium Japanese blends"),
    ("Hong Kong", "Hong Kong", "Teakha", "Tea Brand", "https://www.teakha.com", "@teakha", "discover@teakha.com", "Artisan tea atelier; Sheung Wan; curates rare Asian teas incl. Japanese varieties"),
    # Luxury Hotels
    ("Hong Kong", "Hong Kong", "The Peninsula Hong Kong", "Luxury Hotel", "https://www.peninsula.com/en/hong-kong", "@thepeninsulahotels", "phk@peninsula.com", "Legendary 5-star; iconic afternoon tea; Peninsula blend; premium partner"),
    ("Hong Kong", "Hong Kong", "Rosewood Hong Kong", "Luxury Hotel", "https://www.rosewoodhotels.com/en/hong-kong", "@rosewoodhongkong", "rhongkong@rosewoodhotels.com", "Ultra-luxury Victoria Harbour; Holt's Café tea house; premium F&B partner"),
    ("Hong Kong", "Hong Kong", "Four Seasons Hotel Hong Kong", "Luxury Hotel", "https://www.fourseasons.com/hongkong", "@fshongkong", "res.hongkong@fourseasons.com", "5-star luxury; Caprice & Lung King Heen (Michelin); The Lounge afternoon tea"),
    ("Hong Kong", "Hong Kong", "Mandarin Oriental Hong Kong", "Luxury Hotel", "https://www.mandarinoriental.com/hong-kong", "@mohongkong", "mohkg-reservations@mohg.com", "Landmark luxury; Restaurant Amber (3 Michelin stars); premium tea experience"),
    ("Hong Kong", "Hong Kong", "The Ritz-Carlton Hong Kong", "Luxury Hotel", "https://www.ritzcarlton.com/en/hotels/hong-kong", "@ritzcarltonhk", "rc.hkgrz.leads@ritzcarlton.com", "World's highest hotel floors; Tin Lung Heen fine dining; sky-high afternoon tea"),
    ("Hong Kong", "Hong Kong", "The Upper House", "Luxury Hotel", "https://www.upperhouse.com", "@theupperhouse", "stay@upperhouse.com", "Design-led luxury; Café Gray Deluxe; artisan beverage culture"),
    ("Hong Kong", "Hong Kong", "The St. Regis Hong Kong", "Luxury Hotel", "https://www.marriott.com/hotels/travel/hkgxr-the-st-regis-hong-kong", "@stregishongkong", "stregis.hongkong@stregis.com", "Luxury Wan Chai hotel; Drawing Room afternoon tea; premium butler service"),
    ("Hong Kong", "Hong Kong", "InterContinental Grand Stanford", "Luxury Hotel", "https://www.ihg.com/intercontinental/hotels/gb/en/hong-kong/hkgha", "@ichongkong", "stanhkres@ihg.com", "Tsim Sha Tsui waterfront; multiple dining venues; premium tea interest"),
    ("Hong Kong", "Hong Kong", "Langham Hotel Hong Kong", "Luxury Hotel", "https://www.langhamhotels.com/en/the-langham/hong-kong", "@langhamhongkong", "tlhkg@langhamhotels.com", "Tsim Sha Tsui; T'ang Court (3 Michelin); Artesian bar; premium experiences"),
    ("Hong Kong", "Hong Kong", "Cordis Hong Kong", "Luxury Hotel", "https://www.cordishotels.com/en/hong-kong", "@cordishongkong", "info.hkg@cordishotels.com", "Luxury Mong Kok hotel; The Place dining; upscale café culture"),
    # Luxury Restaurants
    ("Hong Kong", "Hong Kong", "Amber – Mandarin Oriental", "Luxury Restaurant", "https://www.mandarinoriental.com/en/hong-kong/the-landmark/eat-and-drink/restaurants/amber", "@amberhongkong", "amber@mohg.com", "3 Michelin stars 2025; Richard Ekkebus; MICHELIN Green Star; tea pairing focus"),
    ("Hong Kong", "Hong Kong", "Caprice – Four Seasons Hotel HK", "Luxury Restaurant", "https://www.fourseasons.com/hongkong/dining/restaurants/caprice", "@capricehongkong", "caprice.hk@fourseasons.com", "3 Michelin stars; French fine dining; harbor views; premium beverage program"),
    ("Hong Kong", "Hong Kong", "Lung King Heen – Four Seasons Hotel HK", "Luxury Restaurant", "https://www.fourseasons.com/hongkong/dining/restaurants/lung_king_heen", "@fourseasonshk", "lungkingheen.hk@fourseasons.com", "First Chinese restaurant with 3 Michelin stars; refined Cantonese; tea pairings"),
    ("Hong Kong", "Hong Kong", "Ta Vie – The Pottinger Hotel", "Luxury Restaurant", "https://www.tavie.com.hk", "@tavie_hk", "info@tavie.com.hk", "3 Michelin stars; French-Japanese fusion; chef Hideaki Sato; seasonal tea menu"),
    ("Hong Kong", "Hong Kong", "Arbor – H Queen's", "Luxury Restaurant", "https://www.arbor.hk", "@arborhongkong", "info@arbor.hk", "2 Michelin stars; European-Japanese cuisine; chef Eric Räty; premium tea pairing"),
    # ===== BANGKOK (30 stores) =====
    # Specialty Cafes
    ("Thailand", "Bangkok", "Roots Coffee", "Specialty Cafe", "https://rootsbkk.com", "@rootsbkk", "hello@rootsbkk.com", "Pioneer Thai specialty coffee since 2013; 73K IG followers; Thai-only beans"),
    ("Thailand", "Bangkok", "Ceresia Coffee Roasters", "Specialty Cafe", "https://ceresiacoffeeroasters.com", "@ceresiacoffee", "info@ceresiacoffeeroasters.com", "Founded 2013; influential BKK roaster; small batch; tea & coffee menu"),
    ("Thailand", "Bangkok", "Brave Roasters", "Specialty Cafe", "https://www.braveroasters.com", "@braveroasters", "info@braveroasters.com", "Siam Discovery & multiple locations; Thai specialty beans; premium menu"),
    ("Thailand", "Bangkok", "Nana Coffee Roasters", "Specialty Cafe", "https://nanacoffeeroasters.com", "@nanacoffeeroasters", "info@nanacoffeeroasters.com", "World Siphonist Champion 2018; National Brewers Cup 2024; hojicha & matcha menu"),
    ("Thailand", "Bangkok", "Ink & Lion Café", "Specialty Cafe", "https://www.facebook.com/inkandlioncafe", "@inkandlioncafe", "inkandlioncafe@gmail.com", "Ekkamai micro-roastery; art-focused space; tea & coffee gear retail"),
    ("Thailand", "Bangkok", "Mother Roaster", "Specialty Cafe", "https://www.motherroaster.coffee", "@motherroaster", "hello@motherroaster.coffee", "Filter-only; Northern Thai beans; elegant café vibe; premium tea interest"),
    ("Thailand", "Bangkok", "Ratio Coffee & Pastry", "Specialty Cafe", "https://www.ratiobkk.com", "@ratiobkk", "hello@ratiobkk.com", "Specialty café; BKK specialty scene fixture; premium pastry & beverage pairing"),
    ("Thailand", "Bangkok", "Kaizen Coffee Company", "Specialty Cafe", "https://www.kaizencoffee.com", "@kaizencoffeeco", "info@kaizencoffee.com", "Japanese-influenced BKK café; name means 'improvement'; Japanese tea affinity"),
    ("Thailand", "Bangkok", "Qraft / Peace Oriental Teahouse", "Specialty Cafe", "https://www.peaceteahouse.com", "@peaceteahouse", "hello@peaceteahouse.com", "Orient tea-focused; Ari & Empire Tower (55F); full tea pairing menu; top partner"),
    ("Thailand", "Bangkok", "Walden Woods Bangkok", "Specialty Cafe", "https://www.instagram.com/waldenwoodsbangkok", "@waldenwoodsbkk", "waldenwoodsbangkok@gmail.com", "Kyoto-inspired minimalist café; slow-brew filter; Japanese aesthetic; ideal fit"),
    # Tea Brands
    ("Thailand", "Bangkok", "Monsoon Tea", "Tea Brand", "https://www.monsoontea.co.th", "@monsoontea", "info@monsoontea.co.th", "Premium Thai tea brand; wild-harvested teas; specialty tea curators; key partner"),
    ("Thailand", "Bangkok", "Punnpreeda Tea", "Tea Brand", "https://www.punnpreeda.com", "@punnpreeda", "info@punnpreeda.com", "Artisan Thai tea brand; curated premium blends; specialty retail"),
    ("Thailand", "Bangkok", "Doi Chaang Coffee & Tea", "Tea Brand", "https://www.doichaangcoffee.com", "@doichaangcoffee", "info@doichaangcoffee.com", "Renowned Thai highland brand; exports globally; tea & coffee portfolio"),
    ("Thailand", "Bangkok", "O-Oh Farm Suanluang", "Tea Brand", "https://www.o-ohfarm.com", "@o_ohfarm", "contact@o-ohfarm.com", "Organic Thai farm-to-table; specialty tea & produce; eco-conscious partner"),
    ("Thailand", "Bangkok", "Greyhound Café", "Tea Brand", "https://www.greyhoundcafe.co.th", "@greyhoundcafe", "info@greyhoundcafe.co.th", "Iconic BKK lifestyle café-brand; 30+ years; curated premium tea menu potential"),
    # Luxury Hotels
    ("Thailand", "Bangkok", "Mandarin Oriental Bangkok", "Luxury Hotel", "https://www.mandarinoriental.com/en/bangkok", "@mo_bangkok", "mobkk-reservations@mohg.com", "#12 World's 50 Best Hotels 2024; Le Normandie (2 Michelin); Author's Lounge tea"),
    ("Thailand", "Bangkok", "Capella Bangkok", "Luxury Hotel", "https://capellahotels.com/en/capella-bangkok", "@capellabangkok", "bangkok@capellahotels.com", "#1 World's 50 Best Hotels 2024; 101 river-view suites; premium tea lounge"),
    ("Thailand", "Bangkok", "Rosewood Bangkok", "Luxury Hotel", "https://www.rosewoodhotels.com/en/bangkok", "@rosewoodbangkok", "rbangkok@rosewoodhotels.com", "Honeycomb-facade luxury; Sense spa; 4 bars & restaurants; premium F&B partner"),
    ("Thailand", "Bangkok", "Four Seasons Hotel Bangkok at Chao Phraya", "Luxury Hotel", "https://www.fourseasons.com/bangkok", "@fsbangkok", "res.bangkok@fourseasons.com", "#14 World's 50 Best Hotels 2024; MICHELIN Key Hotel; prime riverside dining"),
    ("Thailand", "Bangkok", "The Peninsula Bangkok", "Luxury Hotel", "https://www.peninsula.com/en/bangkok", "@thepeninsulahotels", "pbk@peninsula.com", "Iconic riverside luxury; Wat Arun views; 370 suites; premium afternoon tea"),
    ("Thailand", "Bangkok", "Anantara Riverside Bangkok Resort", "Luxury Hotel", "https://www.anantara.com/en/riverside-bangkok", "@anantara_riverside_bangkok", "riverside.bangkok@anantara.com", "Minor Hotels star property; river cruise experiences; premium Thai tea culture"),
    ("Thailand", "Bangkok", "Park Hyatt Bangkok", "Luxury Hotel", "https://www.hyatt.com/park-hyatt/en-US/bkkph-park-hyatt-bangkok", "@parkhyattbangkok", "parkhyatt.bangkok@hyatt.com", "MICHELIN Key Hotel; The Living Room signature afternoon tea; luxury F&B"),
    ("Thailand", "Bangkok", "The Siam Hotel", "Luxury Hotel", "https://www.thesiamhotel.com", "@thesiamhotel", "info@thesiamhotel.com", "Ultra-boutique art deco riverside; private museum vibes; curated tea program"),
    ("Thailand", "Bangkok", "St. Regis Bangkok", "Luxury Hotel", "https://www.marriott.com/en-us/hotels/bkkxr-the-st-regis-bangkok", "@stregisbangkok", "info.bangkok@stregis.com", "Drawing Room afternoon tea; butler service; premium beverage culture"),
    ("Thailand", "Bangkok", "The Sukhothai Bangkok", "Luxury Hotel", "https://www.sukhothai.com", "@thesukhothaihotel", "info_bkk@sukhothai.com", "MICHELIN Key Hotel; ancient Thai-inspired; serene courtyards; premium tea service"),
    # Luxury Restaurants
    ("Thailand", "Bangkok", "Le Normandie – Mandarin Oriental Bangkok", "Luxury Restaurant", "https://www.mandarinoriental.com/en/bangkok/chao-phraya-river/eat-and-drink/restaurants/le-normandie", "@lenormandie.bangkok", "lenormandie.mobkk@mohg.com", "2 Michelin stars; French fine dining since 1958; prestigious BKK institution"),
    ("Thailand", "Bangkok", "Savelberg Thailand", "Luxury Restaurant", "https://www.savelbergth.com", "@savelbergbangkok", "info@savelbergth.com", "1 Michelin star; Dutch chef with 4 Michelin-starred history; French gastronomy"),
    ("Thailand", "Bangkok", "Mezzaluna by lebua", "Luxury Restaurant", "https://www.lebua.com/mezzaluna", "@mezzalunabangkok", "reservations@lebua.com", "2 Michelin stars; French-Japanese; 65th floor State Tower; tasting menu pairing"),
    ("Thailand", "Bangkok", "Nahm – COMO Metropolitan Bangkok", "Luxury Restaurant", "https://www.comohotels.com/en/metropolitanbangkok/dining/nahm", "@nahmbangkok", "nahm.bkk@comohotels.com", "1 Michelin star; chef Pim Techamuanvivit; iconic Thai destination restaurant"),
    ("Thailand", "Bangkok", "Bo.lan", "Luxury Restaurant", "https://www.bolan.co.th", "@bo.lan_essentially_thai", "bookings@bolan.co.th", "Michelin-starred; organic & sustainable Thai dining; artisan beverage pairings"),
    # ===== DUBAI (30 stores) =====
    # Specialty Cafes
    ("UAE", "Dubai", "Nightjar Coffee Roasters", "Specialty Cafe", "https://www.nightjar.coffee", "@nightjar.coffee", "hello@nightjar.coffee", "Born in Dubai; Alserkal Avenue; in-house roastery; artisan cold brew & nitro"),
    ("UAE", "Dubai", "Fuze Caffè", "Specialty Cafe", "https://www.fuzecaffe.ae", "@fuzecaffe", "info@fuzecaffe.ae", "Specialty café Al Habtoor City; single-origin from Kenya, Ethiopia; Dubai Marina"),
    ("UAE", "Dubai", "Nomad Day Bar (by Nightjar Coffee)", "Specialty Cafe", "https://www.nomaddaybar.com", "@nomaddaybar", "hello@nomaddaybar.com", "Specialty all-day café; One Central; co-working space; Nightjar collaboration"),
    ("UAE", "Dubai", "Mokha 1450", "Specialty Cafe", "https://www.mokha1450.com", "@mokha1450", "info@mokha1450.com", "Luxury UAE specialty coffee brand; prominent regional high-end positioning"),
    ("UAE", "Dubai", "Arabian Tea House", "Specialty Cafe", "https://arabianteahouse.com", "@arabianteahouse", "info@arabianteahouse.com", "Al Fahidi district; 100+ tea varieties; authentic courtyard; Japanese tea fit"),
    ("UAE", "Dubai", "Raw Coffee Company", "Specialty Cafe", "https://www.rawcoffeecompany.com", "@rawcoffeecompany", "info@rawcoffeecompany.com", "Dubai pioneer specialty roaster; direct-trade; Al Quoz roastery; premium teas"),
    ("UAE", "Dubai", "The Sum of Us", "Specialty Cafe", "https://www.thesumofusdubai.com", "@thesumofus_dxb", "hello@thesumofusdubai.com", "Award-winning Dubai café; artisan bakery; specialty coffee & tea; DIFC location"),
    ("UAE", "Dubai", "Stomping Grounds Coffee", "Specialty Cafe", "https://www.stompinggrounds.ae", "@stompinggroundsdxb", "info@stompinggrounds.ae", "Specialty café; City Walk & JLT; curated single-origin; premium tea bar"),
    ("UAE", "Dubai", "Armani/Caffè – Dubai Mall", "Specialty Cafe", "https://www.armani.com/en-ae/armani-caffe", "@armanicaffe_dubai", "dubai.caffe@armani.com", "Luxury fashion café; Fashion Avenue Dubai Mall; all-day dining; premium tea menu"),
    ("UAE", "Dubai", "Blue Box Café by Tiffany – Dubai Mall", "Specialty Cafe", "https://www.tiffany.com", "@tiffanyandco", "dubai.bluebox@tiffany.com", "Only Blue Box Café in Middle East; Dubai Mall; Tea at Tiffany's afternoon tea"),
    # Tea Brands
    ("UAE", "Dubai", "The Cha Club by FLTR", "Tea Brand", "https://www.filteruae.com", "@filtercoffeeuae", "hello@filteruae.com", "Matcha & tea specialist; Level Shoes Dubai Mall; hot/cold matcha & Japanese teas"),
    ("UAE", "Dubai", "Fortnum & Mason Dubai", "Tea Brand", "https://www.fortnumandmason.com", "@fortnumandmason", "dubai@fortnumandmason.com", "300-year British tea heritage; Royal Blend; Dubai Mall luxury retail"),
    ("UAE", "Dubai", "TWG Tea Dubai", "Tea Brand", "https://twgtea.com", "@twgtea", "dubai@twgtea.com", "Singapore luxury tea brand; Dubai Mall & Mall of Emirates; 1,000+ blends"),
    ("UAE", "Dubai", "Project Chaiwala", "Tea Brand", "https://www.projectchaiwala.com", "@projectchaiwala", "info@projectchaiwala.com", "Premium chai-focused brand; Emirati-Indian duo; trendy café concept; tea events"),
    ("UAE", "Dubai", "Dilmah Tea UAE", "Tea Brand", "https://www.dilmah.com", "@dilmahtea", "uae@dilmah.com", "Premium Ceylon tea brand; luxury hotel distributor; UAE market leader; partner"),
    # Luxury Hotels
    ("UAE", "Dubai", "Burj Al Arab Jumeirah", "Luxury Hotel", "https://www.jumeirah.com/en/hotels-resorts/dubai/burj-al-arab", "@burjalarab", "BAA.Reservations@jumeirah.com", "World's most iconic 7-star hotel; butler service; premium afternoon tea program"),
    ("UAE", "Dubai", "Atlantis The Royal", "Luxury Hotel", "https://www.atlantis.com/dubai/atlantis-the-royal", "@atlantistheroyaldubai", "reservations@atlantisthepalm.com", "Forbes triple-5 hotel; 90 pools; Nobu & Ossiano; ultra-premium F&B partner"),
    ("UAE", "Dubai", "Four Seasons Resort Dubai at Jumeirah Beach", "Luxury Hotel", "https://www.fourseasons.com/dubaijb", "@fsdubai", "res.dubaijb@fourseasons.com", "Beachfront luxury; tranquil spa; premium dining & afternoon tea program"),
    ("UAE", "Dubai", "The Ritz-Carlton, Dubai", "Luxury Hotel", "https://www.ritzcarlton.com/en/hotels/dubai", "@ritzcarlton_dubai", "rc.dxbrz.leads@ritzcarlton.com", "JBR oceanfront; Ritz Kids & spa; premium lounge tea service; strong partner"),
    ("UAE", "Dubai", "Bvlgari Resort Dubai", "Luxury Hotel", "https://www.bulgarihotels.com/en_US/dubai", "@bulgarihotels", "dubai@bulgarihotels.com", "Private island off Jumeirah Bay; Italian glamour; Michelin-calibre dining"),
    ("UAE", "Dubai", "Mandarin Oriental Jumeira Dubai", "Luxury Hotel", "https://www.mandarinoriental.com/en/dubai/jumeira", "@mojumeiradubai", "mojbd-reservations@mohg.com", "Jumeirah beachfront; celebrity chef restaurants; holistic spa; premium tea service"),
    ("UAE", "Dubai", "The Lana, Dorchester Collection", "Luxury Hotel", "https://www.dorchestercollection.com/dubai/the-lana", "@thelanadubai", "info.tld@dorchestercollection.com", "Opened 2024 near Burj Khalifa; French-inspired luxury; ultra-premium F&B"),
    ("UAE", "Dubai", "Palazzo Versace Dubai", "Luxury Hotel", "https://www.palazzoversace.ae", "@palazzoversacedubai", "reservations@palazzoversace.ae", "Versace-designed luxury; Mosaico Lounge afternoon tea; premium brand alignment"),
    ("UAE", "Dubai", "Jumeirah Marsa Al Arab", "Luxury Hotel", "https://www.jumeirah.com/en/hotels-resorts/dubai/marsa-al-arab", "@jumeirahmarsaalarab", "marsa.reservations@jumeirah.com", "Superyacht-shaped 2024 hotel; Michelin chef dining; beachfront lounge tea"),
    ("UAE", "Dubai", "One&Only The Palm Dubai", "Luxury Hotel", "https://www.oneandonlyresorts.com/the-palm", "@oneandonlythepalmduabi", "reservations.thepalm@oneandonlyresorts.com", "Palm Jumeirah ultra-luxury; Stay by Yannick Alléno (Michelin); premium tea"),
    # Luxury Restaurants
    ("UAE", "Dubai", "Nobu Dubai – Atlantis The Palm", "Luxury Restaurant", "https://noburestaurants.com/dubai/home", "@nobudubai", "dubai@noburestaurants.com", "Michelin-recognized; 22nd floor; Japanese-Peruvian; Black Cod Miso signature"),
    ("UAE", "Dubai", "Zuma Dubai – DIFC", "Luxury Restaurant", "https://www.zumarestaurant.com/zuma-location/dubai", "@zumadubai", "dubai@zumarestaurant.com", "Contemporary Japanese izakaya; DIFC; celebrity hotspot; premium sake & tea"),
    ("UAE", "Dubai", "Ossiano – Atlantis The Palm", "Luxury Restaurant", "https://www.atlantis.com/dubai/atlantis-the-palm/restaurants/ossiano", "@ossianodubai", "ossiano@atlantisthepalm.com", "Michelin-starred underwater restaurant; seafood tasting menu; premium pairing"),
    ("UAE", "Dubai", "Pierchic – Madinat Jumeirah", "Luxury Restaurant", "https://www.jumeirah.com/en/hotels-resorts/dubai/madinat-jumeirah/restaurants-and-bars/pierchic", "@pierchicdubai", "pierchic@jumeirah.com", "Overwater pier restaurant; Arabian Gulf views; Italian seafood; upscale tea"),
    ("UAE", "Dubai", "Nusr-Et Steakhouse Dubai", "Luxury Restaurant", "https://www.nusr-et.com.tr/en/dubai", "@nusr_et", "dubai@nusr-et.com", "Salt Bae's luxury steakhouse; celebrity dining; premium beverage program"),
]

# Color mapping per country
country_colors = {
    "Singapore": "E8F5E9",
    "Hong Kong": "E3F2FD",
    "Thailand": "FFF8E1",
    "UAE": "FCE4EC",
}

category_colors = {
    "Specialty Cafe": "FFFFFF",
    "Tea Brand": "E8F5E9",
    "Luxury Hotel": "E3F2FD",
    "Luxury Restaurant": "FFF3E0",
}

# Write data rows
for row_idx, row_data in enumerate(data, 2):
    country = row_data[0]
    category = row_data[3]
    fill_color = country_colors.get(country, "FFFFFF")
    if row_idx % 2 == 0:
        row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    else:
        row_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    for col_idx, value in enumerate(row_data, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        cell.fill = row_fill
        cell.border = border
        if col_idx in [1, 2, 4]:
            cell.alignment = center_align
        else:
            cell.alignment = left_align
        cell.font = Font(name="Arial", size=10)

# Set column widths
col_widths = [12, 14, 40, 20, 45, 30, 38, 65]
for col_idx, width in enumerate(col_widths, 1):
    ws.column_dimensions[get_column_letter(col_idx)].width = width

# Freeze top row
ws.freeze_panes = "A2"

# Add auto-filter
ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

# Set row height
ws.row_dimensions[1].height = 25
for row in range(2, len(data) + 2):
    ws.row_dimensions[row].height = 40

# Save file
output_path = "/home/user/Claude-Code-Test/hojicha_export_targets.xlsx"
wb.save(output_path)
print(f"Excel file saved: {output_path}")
print(f"Total rows: {len(data)}")

# Count per city
from collections import Counter
cities = [row[1] for row in data]
print("Count per city:", Counter(cities))
