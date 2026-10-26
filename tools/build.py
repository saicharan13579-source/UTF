"""Build the UTF site from src/ + data/.  Writes index.html (GitHub Pages)."""
import json,os
R=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
rd=lambda p:open(os.path.join(R,p),encoding='utf-8').read()
P=json.loads(rd('src/plots_joined.json'))
s=rd('src/index.src.html')
s=(s.replace('__DATA__',rd('src/data.json')).replace('__PLOTS__',json.dumps(P,separators=(',',':')))
    .replace('__MAP__',rd('src/map.json')).replace('__OV__',rd('data/ov.json')).replace('__DAILY__',rd('data/daily.json'))
    .replace('__EDITS__',rd('data/edits.json')).replace('__CONFIG__',rd('data/config.json')).replace('__IMGW__','3573').replace('__IMGH__','2526'))
head='<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"><style>:root{color-scheme:light;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}body{margin:0;font:14px system-ui,sans-serif}img{max-width:100%}[hidden]{display:none!important}</style></head><body>'
open(os.path.join(R,'index.html'),'w',encoding='utf-8').write(head+s+'</body></html>')
print('built',len(s))
