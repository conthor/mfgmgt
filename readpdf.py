import slate

pdf = 'sample.pdf'
with open(pdf, 'rb') as f:
    doc = slate.PDF(f)

for page in doc[:2]:
    print page
