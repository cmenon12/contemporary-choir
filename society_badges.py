"""Quick script to get data on society badges across the Guild."""
import csv
from datetime import datetime

import requests
from bs4 import BeautifulSoup

# The badges
ALL_BADGES = ["Accessibility Badge", "Democracy Badge", "Development Badge",
              "Diversity Badge", "Fundraising and Volunteering Badge",
              "Inclusivity Badge", "Social Badge", "Sustainability Badge",
              "Welfare Badge"]


def main():
    session = requests.Session()

    # Download all societies
    societies = {}
    page_num = 1

    response = session.get(f"https://my.exeterguild.com/i/get?page={page_num}")
    response.raise_for_status()
    while response.json() != []:
        for item in response.json():
            societies[item["name"]] = item["link"]
        page_num += 1
        response = session.get(f"https://my.exeterguild.com/i/get?page={page_num}")
        response.raise_for_status()
    print(f"We found {len(societies)} societies across {page_num - 1} pages.")

    # Get the number of badges each society has
    csv_data = []
    for name, url in societies.items():
        csv_data.append(get_badges_data(session, name, url))

    # Save to a CSV
    filename = f"society-badges-{datetime.now().isoformat().replace(':', '_')}.csv"
    with open(filename, "w", encoding="UTF8", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["Name", "URL"] + [b.replace(" Badge", "") for b in ALL_BADGES] + ["Total"])
        writer.writerows(csv_data)

    print(f"Saved the results to {filename}.")


def get_badges_data(session: requests.Session, society_name: str, society_url: str) -> list:
    """Get the badge data for the society."""

    # Download it
    badges_got = []
    page_response = session.get(society_url)
    page_response.raise_for_status()
    page_soup = BeautifulSoup(page_response.text, "lxml")
    page_badges = page_soup.find_all("a", {"title": [ALL_BADGES]})
    for item in page_badges:
        badges_got.append(item["title"])

    # Save the result
    result = [society_name, society_url]
    for badge in ALL_BADGES:
        result.append(badge in badges_got)
    result.append(len(badges_got))

    print(result)
    return result


if __name__ == "__main__":
    main()
