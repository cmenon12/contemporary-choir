"""Quick script to get data on society badges across the Guild."""

import requests
from bs4 import BeautifulSoup


def main():
    session = requests.Session()

    # Download all societies
    societies = {}
    all_response = session.get("https://www.exeterguild.org/societies/")
    all_response.raise_for_status()
    all_soup = BeautifulSoup(all_response.text, "lxml")
    all_societies = all_soup.find_all("a", {"class": "msl-listingitem-link"})
    print(f"We found {len(all_societies)} societies.")
    if all_societies is None:
        raise Exception("all_societies is None, the structure has changed!")
    for child in all_societies:
        societies[child.text] = f"https://www.exeterguild.org{child['href']}"

    print(societies)


if __name__ == "__main__":
    main()
