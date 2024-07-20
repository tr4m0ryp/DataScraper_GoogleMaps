import googlemaps
import openpyxl
from openpyxl import Workbook
import os
import time
import logging
import argparse

def create_excel_file(data, filename):
    try:
        workbook = Workbook()
        sheet = workbook.active

        # Write headers
        headers = ["Name", "Address", "Email", "Website"]
        sheet.append(headers)

        # Write data
        for item in data:
            sheet.append(item)

        workbook.save(filename)
        logging.info(f"Data has been written to {filename}")
    except Exception as e:
        logging.error(f"Error writing to Excel file: {e}")

def get_places(api_key, query, location, min_count):
    gmaps = googlemaps.Client(key=api_key)
    places = []
    
    try:
        # Search for places with the given query and location
        places_result = gmaps.places(query=query, location=location)

        # Process results
        while places_result and len(places) < min_count:
            for place in places_result['results']:
                name = place.get('name')
                address = place.get('formatted_address')
                email = ""
                website = place.get('website', "")
                
                # Use place_id to get more details
                place_details = gmaps.place(place_id=place['place_id'])
                if place_details.get('result'):
                    email = place_details['result'].get('email', "")
                    website = place_details['result'].get('website', "")
                
                places.append([name, address, email, website])
            
            # Check for next page token to get more results
            if 'next_page_token' in places_result and len(places) < min_count:
                time.sleep(2)  # Delay to handle rate limiting
                places_result = gmaps.places(query=query, location=location, page_token=places_result['next_page_token'])
            else:
                break
    except Exception as e:
        logging.error(f"Error fetching places: {e}")

    return places

def main():
    logging.basicConfig(level=logging.INFO)

    parser = argparse.ArgumentParser(description="Google Maps Places Scraper")
    parser.add_argument('query', type=str, help="What are you looking for")
    parser.add_argument('location', type=str, help="Location (e.g., country, city)")
    parser.add_argument('min_count', type=int, help="Minimum number of places to collect")
    args = parser.parse_args()

    api_key = os.getenv('GOOGLE_MAPS_API_KEY')
    if not api_key:
        logging.error("API key not found. Please set the GOOGLE_MAPS_API_KEY environment variable.")
        return

    query = args.query
    location = args.location
    min_count = args.min_count

    if not query or not location:
        logging.error("Query and location must be provided.")
        return

    if min_count <= 0:
        logging.error("Minimum count must be greater than 0.")
        return

    data = get_places(api_key, query, location, min_count)

    if data:
        filename = f"{query}_in_{location}.xlsx"
        create_excel_file(data, filename)
    else:
        logging.info("No data found.")

if __name__ == "__main__":
    main()
