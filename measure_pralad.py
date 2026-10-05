import json, statistics, time, urllib.request, http.cookiejar
from pathlib import Path

base='http://127.0.0.1:8000'
opener=urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
opener.open(base+'/index.php').read()
opener.open(base+'/index.php', data=b'regno=070101&password=1234&submit=').read()
results={}
for name,route in [('login','/index.php'),('profile','/my-profile.php'),('history','/enroll-history.php')]:
    samples=[]; statuses=[]; lengths=[]
    for _ in range(30):
        t=time.perf_counter(); response=opener.open(base+route); body=response.read(); samples.append((time.perf_counter()-t)*1000); statuses.append(response.status); lengths.append(len(body))
    ordered=sorted(samples)
    results[name]={'route':route,'requests':len(samples),'status_codes':sorted(set(statuses)),'median_ms':round(statistics.median(samples),2),'p95_ms':round(ordered[int(.95*(len(ordered)-1))],2),'max_ms':round(max(samples),2),'min_body_bytes':min(lengths)}
Path('pralad_evidence/performance.json').write_text(json.dumps(results,indent=2))
print(json.dumps(results,indent=2))
