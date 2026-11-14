"""Compress oversized 2400x1800 artboards into 8-bit palette PNGs; record review metrics."""
import os, json, glob
from PIL import Image
from common import DATA, GFX
files=sorted(glob.glob(os.path.join(GFX,'*.png')))
rec=[]; changed=0
for f in files:
    before=os.path.getsize(f)
    is_an = os.path.basename(f).startswith('AN_')
    if before>250*1024 or is_an:
        im=Image.open(f).convert('RGB')
        im.quantize(colors=256,method=Image.MEDIANCUT,dither=Image.NONE if is_an else Image.FLOYDSTEINBERG).save(f,optimize=True,compress_level=9)
        changed+=1
    after=os.path.getsize(f)
    rec.append(dict(name=os.path.basename(f),before=before,after=after,kb=round(after/1024,1),px=Image.open(f).size))
ovs=[]
for fn in ('graphics_overlap.json','analytics_overlap.json'):
    p=os.path.join(DATA,fn)
    if os.path.exists(p): ovs.append(json.load(open(p)))
ov_total=sum(o['overlapTotal'] for o in ovs); ov_files=sum(len(o['overlaps']) for o in ovs)
review=dict(total=len(files),resized=changed,min_kb=min(r['kb'] for r in rec),max_kb=max(r['kb'] for r in rec),
            mean_kb=round(sum(r['kb'] for r in rec)/len(rec),1),
            resolution='2400 x 1800 px @ 300 DPI (deviceScaleFactor 2)',
            text_overlap_pairs=ov_total,text_overlap_files=ov_files,
            visual_review_samples=['AZNS-207_01_time_donut.png','AZNS-207_02_wellbore.png','AZNS-207_04_fourtrack.png',
                                   'AZNS-207_06_reconciliation.png','FIELD_15_crosssection.png','FIELD_16_learning.png',
                                   'AN_13_pdf_token_census.png','AN_16_cost_bridge.png','AN_17_npt_attribution.png','AN_18_rate_card.png'],
            language='English only inside every graphic',depth_axis='inverted (0 m top)')
json.dump(dict(review=review,files=rec),open(os.path.join(DATA,'visual_review.json'),'w'),indent=1)
print(json.dumps(review,indent=1))
