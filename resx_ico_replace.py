import xml.etree.ElementTree as ET
import base64
import os
import sys

class ResxIconUpdater:
    def __init__(self, icon_path):
        self.icon_path = icon_path
        self.encoded_icon = self._get_encoded_icon()

    def _get_encoded_icon(self):
        if not os.path.exists(self.icon_path):
            print(f"Error: Icon file '{self.icon_path}' not found.")
            return None

        with open(self.icon_path, 'rb') as f:
            iconbytes = f.read()
        # Decode bytes to string for XML compatibility
        return base64.b64encode(iconbytes).decode('utf-8')

    def update_resx_file(self, file_path):
        if not self.encoded_icon:
            raise Exception("Error: No encoded icon available. Cannot update.")

        print(f"Checking {file_path}...")
        # Parse the XML file safely using a context manager to ensure it's closed
        with open(file_path, 'r', encoding='utf-8') as f:
            tree = ET.parse(f)
        root = tree.getroot()
        chunk_size = 80

        found = False
        # Replace Icon
        for item in root.iter('data'):
            if item.attrib.get('name') == '$this.Icon':
                found = True
                # Chunk the base64 string into lines of 80 characters for better readability
                chunked_str = '\n        '.join(self.encoded_icon[i:i+chunk_size] for i in range(0, len(self.encoded_icon), chunk_size))
                item.find('value').text = '\n        ' + chunked_str + '\n    '

        if found:
            print(f"  $this.Icon found in {file_path}, updating...")
            # Overwrite the file
            with open(file_path, 'wb') as f:
                f.write(ET.tostring(root, encoding='utf-8', xml_declaration=True))
        else:
            print(f"  $this.Icon not found in {file_path}, adding...")
            # Add a new data element
            new_data = ET.SubElement(root, 'data', name='$this.Icon')
            new_data.set('type', 'System.Drawing.Icon, System.Drawing')
            new_data.set('mimetype', 'application/x-microsoft.net.object.bytearray.base64')
            new_value = ET.SubElement(new_data, 'value')
            chunked_str = '\n        '.join(self.encoded_icon[i:i+chunk_size] for i in range(0, len(self.encoded_icon), chunk_size))
            new_value.text = '\n        ' + chunked_str + '\n    '
            root.append(new_data)
            with open(file_path, 'wb') as f:
                f.write(ET.tostring(root, encoding='utf-8', xml_declaration=True))


    def search_and_update(self, project_dir, target_filenames = {'mainform.resx', 'form1.resx'}):
        print(f"Searching in: {os.path.abspath(project_dir)}")

        found_any = False
        for root, dirs, files in os.walk(project_dir):
            for file in files:
                if file in target_filenames:
                    found_any = True
                    full_path = os.path.join(root, file)
                    self.update_resx_file(full_path)

        if not found_any:
             raise Exception("{} not found in the project directory".format(target_filenames))

 