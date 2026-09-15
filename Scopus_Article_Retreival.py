# -*- coding: utf-8 -*-
"""
Created on Tue Sep 15 16:54:37 2026

@author: RDASH
"""

import os
import time
import pickle
from datetime import datetime

import requests
import pandas as pd
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

API_KEY = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"  #Elsevier API key
query = '((DOCTYPE(ar) OR DOCTYPE(cp)) AND TITLE-ABS-KEY(("methane decomposition" OR "catalytic methane decomposition" OR "nonoxidative methane decomposition" OR "non-oxidative methane decomposition" OR "thermocatalytic methane decomposition" OR "CH4 decomposition" OR "methane pyrolysis" OR "catalytic pyrolysis of methane") AND ("hydrogen production" OR "H2 production" OR "hydrogen generation" OR "carbon formation" OR "solid carbon" OR "carbon nanotubes" OR "carbon nanofibers") AND NOT ("steam reforming" OR "dry reforming" OR "partial oxidation" OR "oxidative coupling" OR "autothermal reforming" OR "CO2 reforming" OR "bi-reforming" OR "tri-reforming" OR "oxidation" OR "oxidative" OR "O2" OR "oxygen" OR "steam" OR "H2O" OR "CO2" OR "syngas")))'

SEARCH_URL = "https://api.elsevier.com/content/search/scopus"
SERIAL_TITLE_ISSN_URL = "https://api.elsevier.com/content/serial/title/issn/{}"

count = 25
pickle_file = "scopus_articles.pkl"

headers = {
    "X-ELS-APIKey": API_KEY,
    "Accept": "application/json",
}

retry = Retry(
    total=5,
    connect=5,
    read=5,
    backoff_factor=1.5,
    status_forcelist=[429, 500, 502, 503, 504],
    allowed_methods=frozenset(["GET"]),
)

session = requests.Session()
session.headers.update(headers)
session.mount("https://", HTTPAdapter(max_retries=retry))
session.mount("http://", HTTPAdapter(max_retries=retry))

publisher_families = {
    "Elsevier": ["Elsevier"],
    "ACS": ["American Chemical Society"],
    "RSC": ["Royal Society of Chemistry"],
    "Nature": ["Springer Nature", "Nature"],
    "Springer": ["Springer"],
    "Wiley": ["Wiley"],
    "Taylor & Francis": ["Taylor & Francis"],
    "MDPI": ["MDPI"],
}

def get_journal_family(publisher_name: str) -> str:
    if not publisher_name:
        return "Other"
    for family, keywords in publisher_families.items():
        for kw in keywords:
            if kw.lower() in publisher_name.lower():
                return family
    return "Other"

publisher_cache = {}

def normalize_issn(value):
    if not value:
        return ""
    return str(value).replace("-", "").strip()

def get_publisher_from_issn(issn: str) -> str:
    issn = normalize_issn(issn)
    if not issn:
        return ""

    if issn in publisher_cache:
        return publisher_cache[issn]

    try:
        r = session.get(
            SERIAL_TITLE_ISSN_URL.format(issn),
            timeout=(10, 60),
        )
        r.raise_for_status()
        data = r.json()

        entry = data.get("serial-metadata-response", {}).get("entry", [{}])[0]
        publisher = entry.get("dc:publisher", "") or ""
    except Exception as e:
        print(f"Publisher lookup failed for ISSN {issn}: {e}")
        publisher = ""

    publisher_cache[issn] = publisher
    return publisher

# Load previously saved articles if they exist
if os.path.exists(pickle_file):
    with open(pickle_file, "rb") as f:
        articles = pickle.load(f)
    start = len(articles)
    print(f"Resuming from start={start}, {len(articles)} articles already collected")
else:
    articles = []
    start = 0

# First request to get total results
if start == 0:
    try:
        params = {"query": query, "count": 1, "start": 0}
        response = session.get(SEARCH_URL, params=params, timeout=(10, 60))
        response.raise_for_status()
        data = response.json()
        total_results = int(data["search-results"].get("opensearch:totalResults", 0))
        print(f"Total articles found: {total_results}")
    except Exception as e:
        raise RuntimeError(
            f"Initial Scopus request failed. This looks like a network/proxy/firewall issue "
            f"or a temporary Elsevier outage: {e}"
        )
else:
    total_results = None

while total_results is None or start < total_results:
    params = {"query": query, "count": count, "start": start}

    try:
        response = session.get(SEARCH_URL, params=params, timeout=(10, 60))
        response.raise_for_status()
        data = response.json()
    except Exception as e:
        print(f"Search request failed at start={start}: {e}")
        break

    entries = data.get("search-results", {}).get("entry", [])
    if not entries:
        print("No more entries returned, stopping.")
        break

    for item in entries:
        doi = item.get("prism:doi", "")
        title = item.get("dc:title", "")
        journal = item.get("prism:publicationName", "")

        issn = item.get("prism:issn", "") or item.get("prism:eIssn", "")
        publisher = get_publisher_from_issn(issn)
        family = get_journal_family(publisher)

        cover_date = item.get("prism:coverDate", "")
        year = ""
        if cover_date:
            try:
                year = datetime.strptime(cover_date, "%Y-%m-%d").year
            except ValueError:
                year = ""

        articles.append({
            "DOI": doi,
            "Title": title,
            "Journal": journal,
            "ISSN": issn,
            "Publisher": publisher,
            "Journal_Family": family,
            "Year": year,
        })

    start += count
    print(f"Collected {len(articles)} articles so far...")

    with open(pickle_file, "wb") as f:
        pickle.dump(articles, f)

    time.sleep(0.2)

df = pd.DataFrame(articles)
df.to_excel("Scopus Article Retrival.xlsx", index=False)
print("Excel file saved as 'Scopus Article Retrival.xlsx'")



