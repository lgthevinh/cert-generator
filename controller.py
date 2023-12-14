import pandas as pd
import os
from datetime import datetime

def create_folder(folder_path):
  if not os.path.exists(folder_path):
    os.makedirs(folder_path)

def create_markdown(title, tags, name, issue_date, reason):
  markdown_template = f"""
  ---
  title: {title}
  date: {datetime.now().strftime("%Y-%m-%d")}
  categories: ["Certification"]
  tags: ["{tags}"]
  slug: "/certs/{tags}/{name}.lower().replace(" ", "-")"
  name: {name}
  reason: "{reason}"
  issue_date: "{issue_date}"
  signature: "/signatures/tuanle.png"
  undersigned: "Tuan Le - Founder MakerViet"
  ---
  """
  return markdown_template

def generate_certificate(title, tags, name, issue_date, reason):
  markdown = create_markdown(title, tags, name, issue_date, reason)
  create_folder(f"./content/{tags}")
  with open(f"./content/{tags}/{name.lower().replace(' ', '-')}.md", "w") as f:
    f.write(markdown)

def main(filepath, sheet='Sheet1'):
  df = pd.read_excel(f'./content/{filepath}', sheet_name=sheet)
  for _, row in df.iterrows():
    generate_certificate(row['title'], row['tags'], row['name'], row['issue_date'], row['reason'])
