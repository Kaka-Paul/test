import urllib.request, re, json
out=[]
for url in ['https://www.who.int/news-room/fact-sheets/detail/depression','https://www.nimh.nih.gov/health/publications/depression']:
 try:
  req=urllib.request.Request(url,headers={'User-Agent':'Mozilla/5.0'})
  html=urllib.request.urlopen(req,timeout=40).read().decode('utf-8')
  html=re.sub(r'<script\b[^>]*>.*?</script>|<style\b[^>]*>.*?</style>','',html,flags=re.S)
  text=re.sub('<[^>]+>',' ',html)
  text=re.sub(r'\s+',' ',text)
  out.append({'url':url,'text':text})
  print(url,len(text))
 except Exception as e: print(type(e).__name__,str(e))
open('D:/Codex-project/英语第6小组汇报/.ppt_build/sources.json','w',encoding='utf-8').write(json.dumps(out,ensure_ascii=False))
