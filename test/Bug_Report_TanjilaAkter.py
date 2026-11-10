#!/usr/bin/env python3
"""
Bug_Report_TanjilaAkter.py
QA Bug Report Generator + Interactive Dashboard
Target        : Supporters Frame - https://supportersframe.com/
Outputs       : Bug_Report_TanjilaAkter.xlsx  
Usage         : python Bug_Report_TanjilaAkter.py            (console summary + files + opens dashboard)
                python Bug_Report_TanjilaAkter.py --no-open  (do not open browser)
Build .exe    : pyinstaller --onefile --add-data "screenshots;screenshots" Bug_Report_TanjilaAkter.py
Requires      : pip install openpyxl pillow
"""
import base64, html, os, sys, webbrowser
from collections import Counter
from datetime import date

TESTER = "TanjilaAkter"
SITE = "Supporters Frame (https://supportersframe.com/)"
TEST_DATE = date.today().strftime("%d %b %Y")
ENV = "Browsing wbsite of Home, Pricing, Creator, Campaign, Side bar, Login and Sign Up pages (desktop Chrome)"

def res(path):
    """Resolve bundled resource path (works for .py and PyInstaller .exe)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, path)

def out(path):
    """Output files are saved next to the script / exe."""
    base = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, path)

BUGS = [
 dict(id="BUG-001", cat="UI / Layout", title="Hero heading and tagline overlap the frame collage",
      steps="1. Open https://supportersframe.com/ on a phone/tablet-width screen.\n2. Look at the hero section ('show your support').",
      exp="Heading and tagline are fully readable and separated from the collage images.",
      act="The end of the script word 'support(s)' and the last words of the tagline ('...that show') run over the card images, hurting legibility.",
      sev="Medium", shot="BUG01_hero_text_overlaps_collage.png"),
 dict(id="BUG-002", cat="UI / Alignment", title="Trending cards have inconsistent alignment and borders",
      steps="1. Open the home page.\n2. Scroll to 'Trending'.\n3. Compare the five campaign cards.",
      exp="All cards share the same height, border style and title position.",
      act="Cards 1-2 (two-line titles) show titles higher than cards 3-5; card 1 has a heavier border; card borders touch/overlap each other.",
      sev="Medium", shot="BUG02_trending_cards_misaligned.png"),
 dict(id="BUG-003", cat="Data / Logic", title="Top Creators ranking is not in order of supporter count",
      steps="1. Open the home page.\n2. Scroll to 'Top Creators'.\n3. Compare rank numbers with supporter counts.",
      exp="Ranking is consistent with the metric shown (or the ranking criteria are stated).",
      act="#1 Akua Twumasi has 218 supporters while #3 has 426 and #9 'KIM SUNOO LAB 0624' has 1065. Ranking looks inconsistent with displayed numbers.",
      sev="Medium", shot="BUG03_top_creators_ranking_order.png"),
 dict(id="BUG-004", cat="UI / Typography", title="Supporter-count text in Top Creators is very small and inconsistent",
      steps="1. Scroll to 'Top Creators'.\n2. Compare '218 Supporters' with the Trending section.",
      exp="Consistent, readable font size for supporter counts across sections.",
      act="Counts under creator names are tiny and light grey; the same info in Trending is larger.",
      sev="Low", shot="BUG03_top_creators_ranking_order.png"),
 dict(id="BUG-005", cat="UI / Image", title="Creator avatar for 'NEPICIN Support' is cropped",
      steps="1. Scroll to 'Top Creators'.\n2. Look at the #5 avatar.",
      exp="Logo is fully visible inside the circular avatar.",
      act="Logo text is cut off (partial words such as 'Netw...', 'Impl...' visible), so the logo looks broken.",
      sev="Low", shot="BUG03_top_creators_ranking_order.png"),
 dict(id="BUG-006", cat="UI / Text Clipping", title="'Try it now, Free!' link text looks clipped and styled inconsistently",
      steps="1. Scroll to the dark 'gather your supporters' section.\n2. Look beside the 'Learn More' button.",
      exp="Both calls to action use a consistent, fully visible style.",
      act="'Try it now, Free!' is tiny bold text whose last characters appear cut off, unlike the 'Learn More' button.",
      sev="Low", shot="BUG05_cta_text_clipped_and_copy.png"),
 dict(id="BUG-007", cat="Content / Copy", title="Awkward, repetitive headline 'gather your supporters with Supporters'",
      steps="1. Scroll to the dark banner section.\n2. Read the main headline.",
      exp="Clear, grammatical headline (e.g. 'Gather your supporters with Supporters Frame').",
      act="Headline reads 'gather your supporters with Supporters', which is repetitive and unclear.",
      sev="Low", shot="BUG05_cta_text_clipped_and_copy.png"),
 dict(id="BUG-008", cat="UI / Images", title="Partner logo strip is inconsistent and partly cropped",
      steps="1. Scroll to the dark section.\n2. Look at the logo row (WHO, UNICEF, etc.).",
      exp="All logos have uniform tiles/sizing and are fully recognisable.",
      act="Some logos are cropped or faint (2nd and 4th are hard to identify), TEER sits on a dark tile while others are on white, and spacing is uneven.",
      sev="Medium", shot="BUG06_partner_logo_strip.png"),
 dict(id="BUG-009", cat="Accessibility / UI", title="'Need Something Extra?' banner is a flat image with tiny, low-contrast text",
      steps="1. Scroll below the Trending section.\n2. Try to select the banner text or zoom in.",
      exp="Banner text is real, readable, responsive text with a clear button.",
      act="All text and the 'UPGRADE' button are baked into one image; small light-on-light text is hard to read on small screens and not accessible to screen readers.",
      sev="Low", shot="BUG04_upgrade_banner_text_in_image.png"),
 dict(id="BUG-010", cat="UI / Color", title="Footer link 'Bangla Puzzle Limited' uses default blue; copyright text low contrast",
      steps="1. Scroll to the bottom footer.\n2. Read the copyright line.",
      exp="Link color matches site theme; text meets contrast requirements.",
      act="Link appears in default browser blue, unlike the rest of the grey footer; the grey text is very light.",
      sev="Low", shot="BUG07_footer_copyright_link_color.png"),
 dict(id="BUG-011", cat="Text Error", title="Footer label 'Terms & Condition' should be 'Terms & Conditions'",
      steps="1. Scroll to the bottom-right of the footer.",
      exp="'Terms & Conditions'.", act="Label reads 'Terms & Condition'.",
      sev="Low", shot="BUG07_footer_copyright_link_color.png"),
 dict(id="BUG-012", cat="UI / Overlap", title="Floating WhatsApp/Messenger buttons cover hero content",
      steps="1. Open the home page at the top.\n2. Look at the right edge of the hero.",
      exp="Floating buttons do not hide page content.",
      act="The two round buttons sit on top of the collage cards at the right edge.",
      sev="Low", shot="BUG08_floating_chat_buttons_overlap.png"),
 dict(id="BUG-013", cat="Content / Consistency", title="Inconsistent capitalization and labels",
      steps="1. Compare hero buttons ('+ start campaign' vs 'Create Campaign' / 'How It Works').\n2. Compare 'Explore More' and 'View All'.\n3. Check creator names ('DERRICK JOHN JIMENEZ', 'KIM SUNOO LAB 0624').",
      exp="Consistent title case and consistent wording for similar links.",
      act="Mixed lowercase, title case and ALL CAPS; similar 'see more' links use different wording.",
      sev="Low", shot="BUG09_inconsistent_capitalization.png"),
 dict(id="BUG-014", cat="UI / Typography", title="'How It Works' step 3 title shows a smaller '&' and body text is very small",
      steps="1. Scroll to 'How It Works'.\n2. Read step 3 'Customize & Share'.",
      exp="Uniform font size in titles; comfortable body text size.",
      act="The '&' renders visibly smaller than the title words; step descriptions are small and light grey.",
      sev="Low", shot="BUG10_tiny_low_contrast_text.png"),
 dict(id="BUG-015", cat="Text Error", title="Typo in package name: 'Dimond Package' (should be 'Diamond')",
      steps="1. On the home page, click the 'Need Something Extra?' banner.\n2. The 'Special Design Package' (pricing) page opens.\n3. Look at the caption on the Diamond package banner.",
      exp="Caption reads 'Diamond Package' (matching the 'Diamond' title in the image).",
      act="Image title says 'Diamond' but the caption below it says 'Dimond Package'.",
      sev="Medium", shot="BUG15_diamond_typo.png"),
 dict(id="BUG-016", cat="Content / Pricing", title="Package pricing looks illogical: Diamond ($10) equals Gold and is a quarter of Elite ($40)",
      steps="1. On the home page, click the 'Need Something Extra?' banner.\n2. The 'Special Design Package' (pricing) page opens.\n3. Compare prices and features of Gold, Diamond and Elite cards.",
      exp="Price increases consistently with the package tier and its features, or the pricing logic is explained.",
      act="Diamond (1 frame + website integration support) costs 10 $, the same as Gold, while Elite (similar features) costs 40 $. Possible data-entry error; needs confirmation from the product owner.",
      sev="High", shot="BUG16_pricing_diamond_vs_elite.png"),
 dict(id="BUG-018", cat="UI / Images", title="Package banners have duplicated captions and overlapping text",
      steps="1. On the home page, click the 'Need Something Extra?' banner.\n2. The 'Special Design Package' (pricing) page opens.\n3. Compare the large title in each banner with the small caption below it (Silver, Gold, Platinum, Monthly Package, Elite).",
      exp="Each banner shows its package name once, cleanly laid out.",
      act="The package name appears twice (big ribbon text and small white caption). In 'Monthly Package' the ribbon text wraps onto two lines and nearly collides with the caption; logo size/position also differs between cards (e.g. Diamond).",
      sev="Medium", shot="BUG18_banner_caption_duplicates.png"),
 dict(id="BUG-019", cat="Content / Consistency", title="Delivery time is shown in inconsistent formats and values across packages",
      steps="1. On the home page, click the 'Need Something Extra?' banner.\n2. The 'Special Design Package' (pricing) page opens.\n3. Compare the delivery-time line of each package.",
      exp="One consistent format (e.g. '48 hours') and logical timing per tier.",
      act="Silver '2 x 24 Hours', Gold '1 x 24 Hours', Platinum '2 Hours', Diamond '2 Hours', Elite '48 Hours'. Format is mixed, and the higher-priced Elite is slower than cheaper packages.",
      sev="Medium", shot="BUG19_delivery_time_inconsistent.png"),
 dict(id="BUG-020", cat="UI / Content", title="'Revision not applicable' is shown with a green check mark",
      steps="1. On the home page, click the 'Need Something Extra?' banner.\n2. The 'Special Design Package' (pricing) page opens.\n3. Look at the feature list of the Diamond package.",
      exp="Unavailable features are not shown with a positive check icon (or are worded clearly, e.g. 'No revisions').",
      act="'Revision not applicable' has the same green check icon as included features, which is misleading.",
      sev="Low", shot="BUG20_revision_not_applicable.png"),
 dict(id="BUG-021", cat="Content / Pricing", title="Price format '5 $' and no billing period/currency clarity for Monthly Package",
      steps="1. On the home page, click the 'Need Something Extra?' banner.\n2. The 'Special Design Package' (pricing) page opens.\n3. Look at the price on each card, especially 'Monthly Package' (150 $).",
      exp="Standard format such as '$5' or 'USD 5', and a billing period for the monthly plan (e.g. '$150 / month').",
      act="Prices appear as '5 $', '150 $' with no currency code or '/month' label; it is unclear whether the monthly package is recurring or one-time.",
      sev="Low", shot="BUG21_price_format_no_period.png"),
 dict(id="BUG-022", cat="Content / Copy", title="Same generic description copied across several packages",
      steps="1. On the home page, click the 'Need Something Extra?' banner.\n2. The 'Special Design Package' (pricing) page opens.\n3. Read the description text under each price.",
      exp="Description reflects what is special about each package.",
      act="Silver, Gold, Platinum and Monthly use the identical text 'No need to worry about design...'; Diamond and Elite share another identical text. Users must read the bullet list to see differences.",
      sev="Low", shot="BUG22_repeated_description.png"),
 dict(id="BUG-023", cat="UI / Visual", title="Package medal icons are not visually distinct (Silver vs Platinum; Gold, Monthly, Elite)",
      steps="1. On the home page, click the 'Need Something Extra?' banner.\n2. The 'Special Design Package' (pricing) page opens.\n3. Compare the medal icons of all six cards.",
      exp="Each tier has a clearly different color/icon.",
      act="Platinum medal looks silver like the Silver one; Gold, Monthly Package and Elite all use similar gold medals, making tiers hard to tell apart.",
      sev="Low", shot="BUG18_banner_caption_duplicates.png"),
 dict(id="BUG-024", cat="Content / Consistency", title="Brand name written as 'SupportersFrame' in the page subtitle",
      steps="1. On the home page, click the 'Need Something Extra?' banner.\n2. The 'Special Design Package' (pricing) page opens.\n3. Read the subtitle under 'Special Design Package'.",
      exp="Brand written consistently as 'Supporters Frame' (as in logo and footer).",
      act="Subtitle says '...packages from SupportersFrame' (no space).",
      sev="Low", shot="BUG24_brand_name_subtitle.png"),
 dict(id="BUG-025", cat="UI / Spacing", title="Large blank gap between the package cards and the footer",
      steps="1. On the home page, click the 'Need Something Extra?' banner.\n2. The 'Special Design Package' (pricing) page opens.\n3. Scroll below the last row of cards.",
      exp="Consistent spacing before the footer.",
      act="About 150-200 px of empty space is shown below the cards before the footer, unlike spacing used elsewhere.",
      sev="Low", shot="BUG25_layout_blank_gap.png"),
 dict(id="x", cat="Data Consistency", title="Supporter count differs between pages (248 vs 251 vs 253)",
      steps="1. Open the home page and scroll to 'Top Creators'.\n2. Click the creator 'Pixel Sensei For Life' (creator profile opens).\n3. Click the campaign card 'Semaine Culturelle... (J'y Serai)' (campaign page opens).\n4. Compare the supporter count on Top Creators/Trending (248), the creator profile and its campaign card (251), and the campaign page (253 Total Supporters).",
      exp="The same supporter count is shown everywhere (or all pages refresh at the same time).",
      act="Home shows 248, creator profile/card shows 251, campaign page shows 253 for the same creator/campaign. Counts appear stale/cached on some pages.",
      sev="Medium", shot="BUG26_supporter_count_mismatch.png"),
 dict(id="x", cat="UI / Empty State", title="'More Campaigns' section is empty with no cards or message",
      steps="1. Open the home page and scroll to 'Top Creators'.\n2. Click the creator 'Pixel Sensei For Life' (creator profile opens).\n3. Click the campaign card 'Semaine Culturelle... (J'y Serai)' (campaign page opens).\n4. Scroll to the bottom of the campaign page, below the 'Need Something Extra?' banner.",
      exp="Other campaigns are listed, or a message such as 'No more campaigns' is shown, or the section is hidden.",
      act="Heading 'More Campaigns' is shown followed by blank space; no cards, loader or empty-state text.",
      sev="Medium", shot="BUG27_more_campaigns_empty.png"),
 dict(id="x", cat="UI / Clipping", title="'Become Our Partner' button in the top promo banner is cut off",
      steps="1. Open the home page and scroll to 'Top Creators'.\n2. Click the creator 'Pixel Sensei For Life' (creator profile opens).\n3. Click the campaign card 'Semaine Culturelle... (J'y Serai)' (campaign page opens).\n4. Look at the Bangla Puzzle Limited promo banner at the top of the campaign page.",
      exp="The call-to-action button is fully visible inside the banner.",
      act="The 'Become Our Partner' button is cropped at the bottom edge of the banner (only the upper half is visible).",
      sev="Medium", shot="BUG28_partner_banner_cta_clipped.png"),
 dict(id="x", cat="Content / Format", title="'Joined Since' shows a raw timestamp (2026-10-27 13:52:52)",
      steps="1. Open the home page and scroll to 'Top Creators'.\n2. Click the creator 'Pixel Sensei For Life' (creator profile opens).\n3. Look at the stats box on the right of the profile header.",
      exp="User-friendly date such as '27 Oct, 2026' (same format as 'Created on 27 Oct, 2026' on the campaign page).",
      act="Raw technical timestamp with seconds is displayed; the date format differs from the campaign page.",
      sev="Low", shot="BUG29_joined_since_raw_timestamp.png"),
 dict(id="x", cat="UI / Overlap", title="Floating Messenger button overlaps the profile stats box; stats box touches the right edge",
      steps="1. Open the home page and scroll to 'Top Creators'.\n2. Click the creator 'Pixel Sensei For Life' (creator profile opens).\n3. Look at the top-right area where the stats box begins.",
      exp="Floating buttons do not cover content; the box keeps the same side margin as the rest of the page.",
      act="The blue Messenger button overlaps the top of the stats box, and the box runs flush to the right edge of the page with no margin.",
      sev="Low", shot="BUG30_messenger_overlaps_stats_card.png"),
 dict(id="x", cat="UI / Consistency", title="Campaign card on creator profile shows '251' without the 'Supporters' label",
      steps="1. Open the home page and scroll to 'Top Creators'.\n2. Click the creator 'Pixel Sensei For Life' (creator profile opens).\n3. Look at the supporter count under the campaign card.",
      exp="Count is labelled consistently, e.g. '251 Supporters', as in Trending.",
      act="Only an icon and the number '251' are shown, in very small text.",
      sev="Low", shot="BUG31_card_count_no_label.png"),
 dict(id="x", cat="Content / Copy", title="'About This Campaign' text just repeats the campaign title",
      steps="1. Open the home page and scroll to 'Top Creators'.\n2. Click the creator 'Pixel Sensei For Life' (creator profile opens).\n3. Click the campaign card 'Semaine Culturelle... (J'y Serai)' (campaign page opens).\n4. Compare the title with the text under 'About This Campaign'.",
      exp="Description gives details of the campaign (what, when, how to join).",
      act="The description is identical to the title: 'Semaine Culturelle de la Commune du Premier Arrondissement de N'Djaména'.",
      sev="Low", shot="BUG32_description_duplicates_title.png"),
 dict(id="x", cat="Accessibility", title="Icon-only buttons have no visible label (search/filter on profile, header icons)",
      steps="1. Open the home page and scroll to 'Top Creators'.\n2. Click the creator 'Pixel Sensei For Life' (creator profile opens).\n3. Look at the header icons and the two round icons at the right of the 'Campaign' tab.",
      exp="Buttons have a visible label or tooltip so their purpose is clear.",
      act="Only unlabeled icons are shown. Tooltips/aria-labels should be verified manually.",
      sev="Low", shot="BUG33_icon_only_buttons.png"),
 dict(id="x", cat="UI / Spacing", title="Large empty gap in the frame editor panel above 'Choose Photo'",
      steps="1. Open the home page and scroll to 'Top Creators'.\n2. Click the creator 'Pixel Sensei For Life' (creator profile opens).\n3. Click the campaign card 'Semaine Culturelle... (J'y Serai)' (campaign page opens).\n4. Look at the right-hand panel next to the frame preview.",
      exp="Tools and buttons are laid out compactly, or the empty area shows guidance (e.g. 'Choose a photo to start').",
      act="About 200 px of blank space sits between the Filter/Text tools and the Choose Photo/Save buttons.",
      sev="Low", shot="BUG34_editor_panel_empty_gap.png"),
 dict(id="x", cat="UI / Color", title="Stat numbers use different colors (253 blue, 5 green)",
      steps="1. Open the home page and scroll to 'Top Creators'.\n2. Click the creator 'Pixel Sensei For Life' (creator profile opens).\n3. Click the campaign card 'Semaine Culturelle... (J'y Serai)' (campaign page opens).\n4. Look at the stats bar below the 'About This Campaign' section.",
      exp="Stat numbers share one consistent style.",
      act="'253' is blue while '5' is green, with no visible meaning for the color difference.",
      sev="Low", shot="BUG35_stat_number_colors.png"),
 dict(id="x", **{'cat': 'Content / Branding', 'title': "Login popup says 'PhotoFrame' and Sign Up popup says 'Supporter Frame' instead of 'Supporters Frame'", 'steps': "1. Open the home page and click the hamburger menu (top right) to open the side bar.\n2. Click 'Log in' to open the Login popup, or click 'Sign Up' to open the Create Account popup.\n3. Read the sub-title under 'Welcome Back' and under 'Create Account'.", 'exp': "Brand name is written consistently as 'Supporters Frame'.", 'act': "Login popup: 'Login to continue to PhotoFrame' (a different brand name). Sign Up popup: 'Join Supporter Frame today' (singular).", 'sev': 'Medium', 'shot': 'BUG36_brand_name_photoframe.png'}),
 dict(id="x", **{'cat': 'Functional / UI', 'title': 'Side bar stays open behind the Sign Up popup but closes for the Login popup', 'steps': "1. Open the side bar from the hamburger menu.\n2. Click 'Sign Up' and watch the page behind the popup.\n3. Repeat with 'Log in'.", 'exp': 'Side bar closes (or is covered by the dark overlay) the same way for both popups.', 'act': "With 'Sign Up' the white side bar (Explore, Leaderboard, Log in) remains visible and bright beside the popup, outside the dimmed overlay. With 'Log in' the side bar is closed.", 'sev': 'Medium', 'shot': 'BUG38_sidebar_stays_open_behind_signup.png'}),
 dict(id="x", **{'cat': 'Content / Copy', 'title': "Sign Up popup divider says 'or log in with email'", 'steps': "1. Open the home page and click the hamburger menu (top right) to open the side bar.\n2. Click 'Log in' to open the Login popup, or click 'Sign Up' to open the Create Account popup.\n3. Click 'Sign Up' and read the text below 'Sign up with Google'.", 'exp': "'or sign up with email'.", 'act': "Text reads 'or log in with email', which is the Login wording and is confusing on the registration form.", 'sev': 'Low', 'shot': 'BUG37_signup_divider_says_log_in.png'}),
 dict(id="x", **{'cat': 'Content / Validation hint', 'title': "Password placeholders are inconsistent and wrong ('Min 6 Digit' vs 'Minimum 6 characters')", 'steps': "1. Open the home page and click the hamburger menu (top right) to open the side bar.\n2. Click 'Log in' to open the Login popup, or click 'Sign Up' to open the Create Account popup.\n3. Compare the password placeholders in the Login popup and in the Sign Up popup.", 'exp': "Login shows just 'Password'; Sign Up shows one consistent rule such as 'Minimum 6 characters' for both password fields.", 'act': "Login: 'Password Min 6 Digit' (rule shown on a login form). Sign Up: 'Password (Minimum 6 characters)' but 'Confirm Password Min 6 Digit'. 'Digit' wrongly suggests numbers only.", 'sev': 'Low', 'shot': 'BUG39_password_placeholder_wording.png'}),
 dict(id="x", **{'cat': 'UI / Consistency', 'title': 'Login password field has no show/hide (eye) icon, unlike Sign Up', 'steps': "1. Open the home page and click the hamburger menu (top right) to open the side bar.\n2. Click 'Log in' to open the Login popup, or click 'Sign Up' to open the Create Account popup.\n3. Compare the password fields of both popups.", 'exp': 'Both popups give the same option to show/hide the password.', 'act': 'Sign Up password fields have an eye icon; the Login password field has none.', 'sev': 'Low', 'shot': 'BUG40_login_no_eye_icon.png'}),
 dict(id="x", **{'cat': 'Content / Consistency', 'title': "Inconsistent labels: 'Log in' vs 'Login' and 'Sign Up' vs 'Sign up'", 'steps': "1. Open the home page and click the hamburger menu (top right) to open the side bar.\n2. Click 'Log in' to open the Login popup, or click 'Sign Up' to open the Create Account popup.\n3. Compare the side bar buttons with the popup buttons and links.", 'exp': "One consistent spelling and capitalization (e.g. 'Log in' / 'Sign up' everywhere).", 'act': "Side bar uses 'Log in' and 'Sign Up'; popups use 'Login' button, 'Login' link, 'Sign up' link and 'Sign Up' button.", 'sev': 'Low', 'shot': 'BUG41_login_signup_label_inconsistent.png'}),
 dict(id="x", **{'cat': 'UI / Typography', 'title': "Question marks and apostrophes render tiny in popup text ('Password?', 'account?', 'Don't')", 'steps': "1. Open the home page and click the hamburger menu (top right) to open the side bar.\n2. Click 'Log in' to open the Login popup, or click 'Sign Up' to open the Create Account popup.\n3. Look at 'Forgot Password?' and 'Don't have an account?' in the Login popup (also 'Already have an account?' in Sign Up).", 'exp': 'Punctuation is the same size and style as the letters.', 'act': "The question mark and apostrophe appear as small raised glyphs, which looks like a font fallback problem. The same effect is seen in 'Today's Supporters' and 'N'Djaména' on the campaign page.", 'sev': 'Low', 'shot': 'BUG42_punctuation_glyph_small.png'}),
 dict(id="x", **{'cat': 'UI / Design', 'title': 'Side bar: red close icon is off-theme and Log in / Sign Up buttons have identical style', 'steps': '1. Open the side bar from the hamburger menu.\n2. Look at the close (X) icon and the two buttons at the bottom.', 'exp': 'Close icon matches the site color scheme; the main action (Sign Up) is visually emphasised.', 'act': 'Red X does not match the teal/grey theme; both buttons are plain outline buttons with the same weight, so there is no clear primary action; button text is larger than the menu items.', 'sev': 'Low', 'shot': 'BUG43_sidebar_close_red_buttons_same.png'}),
]
for _i, _b in enumerate(BUGS, 1):
    _b["id"] = f"BUG-{_i:03d}"  # auto-number so IDs stay sequential when you add/remove bugs

COVERAGE = [
 ("UI/UX - Home page (alignment, spacing, color, typography)", "Reviewed from screenshot", "Done"),
 ("UI/UX + content - Pricing page (banner click)", "Reviewed from screenshot", "Done"),
 ("Order Now buttons (open WhatsApp?)", "Click each button on pricing page", "TO DO"),
 ("UI/UX - Creator profile + Campaign page", "Reviewed from screenshots", "Done"),
 ("Frame editor: Choose Photo, Filter, Text, Save, Download QR, share link", "Upload a photo and test each control", "TO DO"),
 ("UI/UX - Side bar, Login and Sign Up popups", "Reviewed from screenshots", "Done"),
 ("Login/Sign Up validation (empty, invalid email, short password, mismatch confirm, duplicate email)", "Test each case manually", "TO DO"),
 ("Google login, Forgot Password, show/hide password, popup close (X / Esc / click outside)", "Test manually", "TO DO"),
 ("Screenshots taken with browser extensions OFF", "Use Incognito to avoid extension overlays", "TO DO"),
 ("Text errors / copy", "Reviewed from screenshot", "Done"),
 ("Broken links and 404 pages", "Click every menu/footer link", "TO DO"),
 ("Functional (forms, buttons, search, login)", "Test manually", "TO DO"),
 ("Mobile responsiveness (375 / 768 / 1366 px)", "DevTools device mode", "TO DO"),
 ("Performance / console errors", "F12 Console, Network, Lighthouse", "TO DO"),
]
SEV_ORDER = {"High": 0, "Medium": 1, "Low": 2}
SEV_COLOR = {"High": "C0392B", "Medium": "E67E22", "Low": "27AE60"}

# ------------------------------------------------------------------ Excel
def export_excel(path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from PIL import Image as PILImage
    wb = Workbook(); ws = wb.active; ws.title = "Bug Report"
    f = "Arial"; thin = Side(style="thin", color="BBBBBB"); border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws["A1"] = f"Bug Report - {SITE}"; ws["A1"].font = Font(name=f, size=16, bold=True, color="1F2A44")
    ws["A2"] = f"Tester: {TESTER}   |   Date: {TEST_DATE}   |   Environment: {ENV}"; ws["A2"].font = Font(name=f, size=10, italic=True)
    heads = ["Bug ID", "Category", "Bug Title / Summary", "Steps to Reproduce", "Expected Result", "Actual Result", "Severity", "Status", "Screenshot"]
    hr = 4
    for c, h in enumerate(heads, 1):
        cell = ws.cell(hr, c, h); cell.font = Font(name=f, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F2A44"); cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); cell.border = border
    widths = [10, 18, 36, 48, 36, 48, 11, 10, 34]
    for i, w in enumerate(widths): ws.column_dimensions[chr(65 + i)].width = w
    r = hr + 1
    for b in BUGS:
        vals = [b["id"], b["cat"], b["title"], b["steps"], b["exp"], b["act"], b["sev"], "Open", ""]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v); cell.font = Font(name=f, size=10); cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="center" if c in (1, 7, 8) else "left")
        sc = ws.cell(r, 7); sc.font = Font(name=f, size=10, bold=True, color="FFFFFF"); sc.fill = PatternFill("solid", fgColor=SEV_COLOR[b["sev"]])
        ws.row_dimensions[r].height = 80
        p = res(os.path.join("screenshots", b["shot"])) if b["shot"] else ""
        if p and os.path.exists(p):
            im = PILImage.open(p); w, h = im.size; s = min(230 / w, 100 / h, 1)
            img = XLImage(p); img.width, img.height = int(w * s), int(h * s)
            ws.row_dimensions[r].height = max(80, img.height * 0.75 + 6); ws.add_image(img, f"I{r}")
        elif not b["shot"]:
            ws.cell(r, 9, "N/A (measured via DevTools / behaviour)").alignment = Alignment(wrap_text=True, vertical="top")
            ws.cell(r, 9).font = Font(name=f, size=9, italic=True)
        r += 1
    last = r - 1
    ws.freeze_panes = ws.cell(hr + 1, 4); ws.auto_filter.ref = f"A{hr}:I{last}"
    # Summary sheet (live formulas)
    s = wb.create_sheet("Summary")
    s["A1"] = "Severity Summary"; s["A1"].font = Font(name=f, size=14, bold=True, color="1F2A44")
    for c, h in enumerate(["Severity", "Count", "% of Total"], 1):
        x = s.cell(3, c, h); x.font = Font(name=f, bold=True, color="FFFFFF"); x.fill = PatternFill("solid", fgColor="1F2A44"); x.alignment = Alignment(horizontal="center")
    for i, sv in enumerate(["High", "Medium", "Low"], 4):
        s.cell(i, 1, sv).font = Font(name=f, bold=True, color="FFFFFF"); s.cell(i, 1).fill = PatternFill("solid", fgColor=SEV_COLOR[sv])
        s.cell(i, 2, f"=COUNTIF('Bug Report'!$G$5:$G${last},A{i})")
        s.cell(i, 3, f"=IF($B$7=0,0,B{i}/$B$7)"); s.cell(i, 3).number_format = "0.0%"
    s["A7"] = "Total"; s["B7"] = "=SUM(B4:B6)"; s["C7"] = "=SUM(C4:C6)"; s["C7"].number_format = "0.0%"
    for rr in range(4, 8):
        for cc in range(2, 4): s.cell(rr, cc).font = Font(name=f, bold=(rr == 7)); s.cell(rr, cc).alignment = Alignment(horizontal="center")
    s["A7"].font = Font(name=f, bold=True)
    s["A9"] = "Category Summary"; s["A9"].font = Font(name=f, size=14, bold=True, color="1F2A44")
    cats = sorted({b["cat"] for b in BUGS})
    for c, h in enumerate(["Category", "Count"], 1):
        x = s.cell(10, c, h); x.font = Font(name=f, bold=True, color="FFFFFF"); x.fill = PatternFill("solid", fgColor="1F2A44")
    for i, ct in enumerate(cats, 11):
        s.cell(i, 1, ct).font = Font(name=f); s.cell(i, 2, f"=COUNTIF('Bug Report'!$B$5:$B${last},A{i})").font = Font(name=f)
    s["A%d" % (12 + len(cats))] = "Note: counts are formulas linked to the 'Bug Report' sheet; update Severity/Category there and this sheet recalculates."
    s["A%d" % (12 + len(cats))].font = Font(name=f, size=9, italic=True)
    s.column_dimensions["A"].width = 26; s.column_dimensions["B"].width = 12; s.column_dimensions["C"].width = 12
    from openpyxl.chart import BarChart, Reference
    ch = BarChart(); ch.title = "Bugs by Severity"; ch.legend = None
    ch.add_data(Reference(s, min_col=2, min_row=3, max_row=6), titles_from_data=True); ch.set_categories(Reference(s, min_col=1, min_row=4, max_row=6))
    ch.height, ch.width = 7, 11; s.add_chart(ch, "E3")
    cv = wb.create_sheet("Test Coverage")
    for c, h in enumerate(["Area", "Method", "Status"], 1):
        x = cv.cell(1, c, h); x.font = Font(name=f, bold=True, color="FFFFFF"); x.fill = PatternFill("solid", fgColor="1F2A44")
    for i, row in enumerate(COVERAGE, 2):
        for c, v in enumerate(row, 1): cv.cell(i, c, v).font = Font(name=f)
    cv.column_dimensions["A"].width = 46; cv.column_dimensions["B"].width = 38; cv.column_dimensions["C"].width = 24
    wb.save(path)

# ------------------------------------------------------------------ HTML dashboard
def b64(p):
    try:
        with open(p, "rb") as fh: return "data:image/png;base64," + base64.b64encode(fh.read()).decode()
    except OSError: return ""

def export_html(path):
    cnt = Counter(b["sev"] for b in BUGS); total = len(BUGS)
    rows = []
    for b in sorted(BUGS, key=lambda x: (SEV_ORDER[x["sev"]], x["id"])):
        img = b64(res(os.path.join("screenshots", b["shot"]))) if b["shot"] else ""
        shot = f'<img src="{img}" alt="screenshot {b["id"]}">' if img else "<em>N/A</em>"
        esc = lambda t: html.escape(t).replace("\n", "<br>")
        rows.append(f'<tr data-sev="{b["sev"]}"><td><b>{b["id"]}</b></td><td>{html.escape(b["cat"])}</td><td><b>{html.escape(b["title"])}</b></td>'
                    f'<td>{esc(b["steps"])}</td><td>{esc(b["exp"])}</td><td>{esc(b["act"])}</td>'
                    f'<td><span class="sev {b["sev"]}">{b["sev"]}</span></td><td>{shot}</td></tr>')
    bars = "".join(f'<div class="bar"><span>{s}</span><div class="fill {s}" style="width:{cnt[s]/total*100:.0f}%">{cnt[s]}</div></div>' for s in ["High", "Medium", "Low"])
    doc = f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>Bug Report - {html.escape(SITE)}</title>
<style>
:root{{--bg:#f4f6fb;--fg:#1f2a44;--card:#fff;--bd:#d8dce6}}
@media(prefers-color-scheme:dark){{:root{{--bg:#12151d;--fg:#e6e9f2;--card:#1b2030;--bd:#2d3447}}}}
body{{font-family:Arial,sans-serif;background:var(--bg);color:var(--fg);margin:0;padding:24px}}
h1{{margin:0}} .meta{{opacity:.75;font-size:13px;margin:6px 0 18px}}
.cards{{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:18px}}
.card{{background:var(--card);border:1px solid var(--bd);border-radius:8px;padding:14px 20px;min-width:130px}}
.card b{{font-size:28px;display:block}}
.bar{{display:flex;align-items:center;gap:8px;margin:4px 0;width:260px}} .bar span{{width:60px;font-size:13px}}
.fill{{color:#fff;font-weight:bold;padding:2px 6px;border-radius:4px;min-width:22px}}
.High{{background:#c0392b}}.Medium{{background:#e67e22}}.Low{{background:#27ae60}}
.tools{{margin:12px 0}} button,input{{padding:7px 12px;border:1px solid var(--bd);border-radius:4px;background:var(--card);color:var(--fg);cursor:pointer}}
button.on{{background:var(--fg);color:var(--bg)}}
.wrap{{overflow-x:auto}} table{{border-collapse:collapse;width:100%;background:var(--card);font-size:13px}}
th{{background:#1f2a44;color:#fff;padding:8px;text-align:left;position:sticky;top:0}} td{{border:1px solid var(--bd);padding:8px;vertical-align:top;min-width:90px}}
td img{{width:200px;border:1px solid var(--bd);cursor:zoom-in}} td img.big{{width:640px;cursor:zoom-out}}
.sev{{color:#fff;font-weight:bold;padding:3px 9px;border-radius:10px}}
</style></head><body>
<h1>Bug Report &mdash; {html.escape(SITE)}</h1>
<div class="meta">Tester: {TESTER} &nbsp;|&nbsp; Date: {TEST_DATE} &nbsp;|&nbsp; {html.escape(ENV)}</div>
<div class="cards"><div class="card"><b>{total}</b>Total bugs</div><div class="card"><b>{cnt['High']}</b>High</div><div class="card"><b>{cnt['Medium']}</b>Medium</div><div class="card"><b>{cnt['Low']}</b>Low</div><div class="card">{bars}</div></div>
<div class="tools"><button class="on" data-f="All">All</button> <button data-f="High">High</button> <button data-f="Medium">Medium</button> <button data-f="Low">Low</button>
 <input id="q" placeholder="Search bugs..."></div>
<div class="wrap"><table id="t"><thead><tr><th>Bug ID</th><th>Category</th><th>Title</th><th>Steps to Reproduce</th><th>Expected</th><th>Actual</th><th>Severity</th><th>Screenshot</th></tr></thead><tbody>{''.join(rows)}</tbody></table></div>
<script>
var f='All',q='';function ap(){{document.querySelectorAll('#t tbody tr').forEach(function(r){{r.style.display=((f=='All'||r.dataset.sev==f)&&r.innerText.toLowerCase().indexOf(q)>-1)?'':'none'}})}}
document.querySelectorAll('button[data-f]').forEach(function(b){{b.onclick=function(){{f=b.dataset.f;document.querySelectorAll('button[data-f]').forEach(function(x){{x.classList.remove('on')}});b.classList.add('on');ap()}}}});
document.getElementById('q').oninput=function(e){{q=e.target.value.toLowerCase();ap()}};
document.querySelectorAll('td img').forEach(function(i){{i.onclick=function(){{i.classList.toggle('big')}}}});
</script></body></html>"""
    with open(path, "w", encoding="utf-8") as fh: fh.write(doc)

# ------------------------------------------------------------------ Console
def console_summary():
    cnt = Counter(b["sev"] for b in BUGS)
    C = {"High": "\033[91m", "Medium": "\033[93m", "Low": "\033[92m"}; R = "\033[0m"
    print("=" * 78); print(f" BUG REPORT  |  {SITE}"); print(f" Tester: {TESTER}  |  Date: {TEST_DATE}"); print("=" * 78)
    for b in sorted(BUGS, key=lambda x: (SEV_ORDER[x["sev"]], x["id"])):
        print(f" {b['id']}  {C[b['sev']]}{b['sev']:<6}{R}  [{b['cat']}] {b['title']}")
    print("-" * 78)
    for s in ["High", "Medium", "Low"]: print(f" {C[s]}{s:<7}{R}: {'#' * cnt[s]} {cnt[s]}")
    print(f" Total  : {len(BUGS)}"); print("=" * 78)

def main():
    os.system("")  # enable ANSI colours on Windows terminals
    console_summary()
    x, h = out(f"Bug_Report_{TESTER}.xlsx"), out(f"Bug_Report_{TESTER}_Dashboard.html")
    export_excel(x); export_html(h)
    print(f" Excel     saved: {x}\n Dashboard saved: {h}")
    if "--no-open" not in sys.argv:
        webbrowser.open("file:///" + h.replace("\\", "/"))
        try: input("\nPress Enter to exit...")
        except EOFError: pass

if __name__ == "__main__":
    main()
