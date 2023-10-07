import json


def remove_null_key(json_obj):
    if isinstance(json_obj, dict):
        # Create a copy of the dictionary to avoid modifying it while iterating
        keys_to_remove = []
        for key, value in json_obj.items():
            if key == "null":
                keys_to_remove.append(key)
            else:
                remove_null_key(value)

        for key in keys_to_remove:
            del json_obj[key]
    elif isinstance(json_obj, list):
        # Recursively remove null keys from elements in the list
        for item in json_obj:
            remove_null_key(item)


def remove_null_keys_from_json_file(file_path):
    try:
        # Read JSON data from the file
        with open(file_path, 'r') as file:
            data = json.load(file)

        # Remove the "null" key and its associated values
        remove_null_key(data)

        # Write the modified JSON data back to the file
        with open(file_path, 'w') as file:
            json.dump(data, file, indent=4)
    except FileNotFoundError:
        print(f"File not found: {file_path}")
    except json.JSONDecodeError:
        print(f"Invalid JSON format in file: {file_path}")
    except Exception as e:
        print(f"An error occurred: {str(e)}")


# Usage example
file_path = 'data.json'
remove_null_keys_from_json_file(file_path)
