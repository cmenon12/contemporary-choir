"""Quick script to get data on society badges across the Guild."""
import csv
from datetime import datetime

import requests
from bs4 import BeautifulSoup

# The badges, as per the text in the a.msl-groupingattributelist-link tags
ALL_BADGES = ["Accessibility", "Democracy", "Development", "Diversity", "Fundraising", "Inclusivity", "Social",
              "Sustainability", "welfare"]


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
    for item in all_societies:
        societies[item.text] = f"https://www.exeterguild.org{item['href']}"

    # Get the number of badges each society has
    csv_data = []
    for name, url in societies.items():
        csv_data.append(get_badges_data(session, name, url))

    # Save to a CSV
    filename = f"society-badges-{datetime.now().isoformat().replace(':', '_')}.csv"
    with open(filename, "w", encoding="UTF8", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["Name", "URL"] + ALL_BADGES + ["Total"])
        writer.writerows(csv_data)

    print(f"Saved the results to {filename}.")


def get_badges_data(session: requests.Session, society_name: str, society_url: str) -> list:
    """Get the badge data for the society."""

    # Download it
    badges_got = []
    page_response = session.get(society_url)
    page_response.raise_for_status()
    page_soup = BeautifulSoup(page_response.text, "lxml")
    page_badges = page_soup.find_all("a", {"class": "msl-groupingattributelist-link"})
    for item in page_badges:
        badges_got.append(item.text)

    # Save the result
    result = [society_name, society_url]
    for badge in ALL_BADGES:
        result.append(badge in badges_got)
    result.append(len(badges_got))

    print(result)
    return result


if __name__ == "__main__":
    main()
