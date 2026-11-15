# Fix audio filenames in infographic.html
with open('infographic.html', 'r', encoding='utf-8') as f:
    info_content = f.read()

info_content = info_content.replace('audio/post_office_en.mp3', 'audio/post_office.mp3')
info_content = info_content.replace('audio/notre_dame_en.mp3', 'audio/notre_dame.mp3')
info_content = info_content.replace('audio/ben_thanh_en.mp3', 'audio/ben_thanh.mp3')
info_content = info_content.replace('audio/dragon_wharf_en.mp3', 'audio/dragon_wharf.mp3')

with open('infographic.html', 'w', encoding='utf-8') as f:
    f.write(info_content)

with open('03_Website_Hoc_Lieu_So/infographic.html', 'w', encoding='utf-8') as f:
    f.write(info_content)

print("Updated infographic.html audio sources")

# Fix QR fallback targetUrl in index.html
with open('index.html', 'r', encoding='utf-8') as f:
    index_content = f.read()

# Replace in pres-qr-zoom
old_zoom = '''      let targetUrl = `infographic.html?site=${site.id}`;
      try {
        const loc = window.location;
        if (loc.protocol !== 'file:') {
          targetUrl = `${loc.origin}${loc.pathname.substring(0, loc.pathname.lastIndexOf('/') + 1)}infographic.html?site=${site.id}`;
        }
      } catch(e) {}'''

new_zoom = '''      let targetUrl = `https://quangthien18041996-stack.github.io/ENGLISH8/infographic.html?site=${site.id}`;
      try {
        const loc = window.location;
        if (loc.protocol !== 'file:' && loc.origin && !loc.origin.includes('null')) {
          targetUrl = `${loc.origin}${loc.pathname.substring(0, loc.pathname.lastIndexOf('/') + 1)}infographic.html?site=${site.id}`;
        }
      } catch(e) {}'''

# Replace in pres-poster-canvas qrBox
old_poster = '''        let targetUrl = `infographic.html?site=${site.id}`;
        try {
          const loc = window.location;
          if (loc.protocol !== 'file:') {
            targetUrl = `${loc.origin}${loc.pathname.substring(0, loc.pathname.lastIndexOf('/') + 1)}infographic.html?site=${site.id}`;
          }
        } catch(e) {}'''

new_poster = '''        let targetUrl = `https://quangthien18041996-stack.github.io/ENGLISH8/infographic.html?site=${site.id}`;
        try {
          const loc = window.location;
          if (loc.protocol !== 'file:' && loc.origin && !loc.origin.includes('null')) {
            targetUrl = `${loc.origin}${loc.pathname.substring(0, loc.pathname.lastIndexOf('/') + 1)}infographic.html?site=${site.id}`;
          }
        } catch(e) {}'''

# Replace in showQrModal
old_modal = '''      let targetUrl = `infographic.html?site=${siteId}`;
      try {
        const loc = window.location;
        if (loc.protocol !== 'file:') {
          targetUrl = `${loc.origin}${loc.pathname.substring(0, loc.pathname.lastIndexOf('/') + 1)}infographic.html?site=${siteId}`;
        }
      } catch(e) {}'''

new_modal = '''      let targetUrl = `https://quangthien18041996-stack.github.io/ENGLISH8/infographic.html?site=${siteId}`;
      try {
        const loc = window.location;
        if (loc.protocol !== 'file:' && loc.origin && !loc.origin.includes('null')) {
          targetUrl = `${loc.origin}${loc.pathname.substring(0, loc.pathname.lastIndexOf('/') + 1)}infographic.html?site=${siteId}`;
        }
      } catch(e) {}'''

if old_zoom in index_content:
    index_content = index_content.replace(old_zoom, new_zoom, 1)
    print("Replaced zoom QR targetUrl")

if old_poster in index_content:
    index_content = index_content.replace(old_poster, new_poster, 1)
    print("Replaced poster QR targetUrl")

if old_modal in index_content:
    index_content = index_content.replace(old_modal, new_modal, 1)
    print("Replaced modal QR targetUrl")

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(index_content)

print("Updated index.html QR target URLs")
