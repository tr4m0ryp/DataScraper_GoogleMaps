import googlemaps
import openpyxl
from openpyxl import Workbook
import os
import time
import logging
from rich.console import Console
from rich.prompt import Prompt
from rich.text import Text

# Initialize rich console
console = Console()

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
        console.log(f"Data has been written to [bold green]{filename}[/bold green]")
    except Exception as e:
        console.log(f"[bold red]Error writing to Excel file:[/bold red] {e}")

def get_places(api_key, query, location, min_count):
    gmaps = googlemaps.Client(key=api_key)
    places = []
    start_time = time.time()
    
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

                # Display progress
                elapsed_time = time.time() - start_time
                avg_time_per_place = elapsed_time / len(places) if places else 0
                estimated_total_time = avg_time_per_place * min_count
                remaining_time = estimated_total_time - elapsed_time
                console.log(f"Collected [bold blue]{len(places)}[/bold blue]/{min_count} places. Estimated time remaining: [bold yellow]{remaining_time:.2f}[/bold yellow] seconds.")

            # Check for next page token to get more results
            if 'next_page_token' in places_result and len(places) < min_count:
                time.sleep(2)  # Delay to handle rate limiting
                places_result = gmaps.places(query=query, location=location, page_token=places_result['next_page_token'])
            else:
                break
    except Exception as e:
        console.log(f"[bold red]Error fetching places:[/bold red] {e}")

    return places

def main():
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    # Animated user input with rich
    query = Prompt.ask(Text("What are you looking for?", style="bold cyan"))
    location = Prompt.ask(Text("Enter the location (e.g., country, city):", style="bold cyan"))
    min_count = Prompt.ask(Text("Minimum number of places to collect:", style="bold cyan"), default="10", show_default=False)
    
    min_count = int(min_count) if min_count.isdigit() else 10

    api_key = os.getenv('GOOGLE_MAPS_API_KEY')
    if not api_key:
        console.log("[bold red]API key not found. Please set the GOOGLE_MAPS_API_KEY environment variable.[/bold red]")
        return

    console.log(f"Starting the scrape for '[bold blue]{query}[/bold blue]' in '[bold blue]{location}[/bold blue]' with a minimum count of [bold blue]{min_count}[/bold blue].")
    data = get_places(api_key, query, location, min_count)

    if data:
        filename = f"{query}_in_{location}.xlsx"
        create_excel_file(data, filename)
    else:
        console.log("[bold yellow]No data found.[/bold yellow]")

if __name__ == "__main__":
    main()
