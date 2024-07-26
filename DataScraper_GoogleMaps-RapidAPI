import requests
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
        if sheet.max_row == 1:
            headers = ["Name", "Address"]
            sheet.append(headers)

        # Write data
        for item in data:
            sheet.append(item)

        workbook.save(filename)
        console.log(f"Data has been written to [bold green]{filename}[/bold green]")
    except Exception as e:
        console.log(f"[bold red]Error writing to Excel file:[/bold red] {e}")

def get_places(api_key, query, location, min_count):
    url = "https://google-maps28.p.rapidapi.com/maps/api/place/textsearch/json"
    headers = {
        "X-RapidAPI-Key": api_key,
        "X-RapidAPI-Host": "google-maps28.p.rapidapi.com"
    }
    params = {
        "query": query,
        "region": location,
        "type": "point_of_interest",
        "radius": 10000
    }
    places = []
    start_time = time.time()
    filename = f"{query}_in_{location}.xlsx"
    result_count = 0

    try:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        places_result = response.json()

        while places_result and len(places) < min_count:
            for place in places_result['results']:
                name = place.get('name')
                address = place.get('formatted_address')
                places.append([name, address])
                result_count += 1

                # Save data every 1000 places
                if result_count % 1000 == 0:
                    create_excel_file(places, filename)
                    places.clear()  # Clear the list after saving to prevent duplicate saving
                    console.log(f"Partial data saved at {result_count} places.")

                # Display progress
                elapsed_time = time.time() - start_time
                avg_time_per_place = elapsed_time / result_count if result_count else 0
                estimated_total_time = avg_time_per_place * min_count
                remaining_time = estimated_total_time - elapsed_time
                console.log(f"Collected [bold blue]{result_count}[/bold blue]/{min_count} places. Estimated time remaining: [bold yellow]{remaining_time:.2f}[/bold yellow] seconds.")

            # Check for next page token to get more results
            if 'next_page_token' in places_result and len(places) < min_count:
                params['pagetoken'] = places_result['next_page_token']
                time.sleep(2)  # Delay to handle rate limiting
                response = requests.get(url, headers=headers, params=params)
                response.raise_for_status()
                places_result = response.json()
            else:
                break
    except requests.exceptions.RequestException as e:
        console.log(f"[bold red]Error fetching places:[/bold red] {e}")

    # Save remaining data
    if places:
        create_excel_file(places, filename)

    return places

def main():
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    # Animated user input with rich
    api_key = Prompt.ask(Text("Enter your Google Maps API key:", style="bold cyan"), password=True)
    query = Prompt.ask(Text("What are you looking for?", style="bold cyan"))
    location = Prompt.ask(Text("Enter the location (e.g., country, city):", style="bold cyan"))
    min_count = Prompt.ask(Text("Minimum number of places to collect:", style="bold cyan"), default="10", show_default=False)
    
    min_count = int(min_count) if min_count.isdigit() else 10

    if not api_key:
        console.log("[bold red]API key not provided. Please enter your Google Maps API key.[/bold red]")
        return

    console.log(f"Starting the scrape for '[bold blue]{query}[/bold blue]' in '[bold blue]{location}[/bold blue]' with a minimum count of [bold blue]{min_count}[/bold blue].")
    data = get_places(api_key, query, location, min_count)

    if data:
        console.log(f"[bold green]Data collection complete.[/bold green]")
    else:
        console.log("[bold yellow]No data found or an error occurred during data collection.[/bold yellow]")

if __name__ == "__main__":
    main()
