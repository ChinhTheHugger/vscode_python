import osmium
import io

class OSMHandler(osmium.SimpleHandler):
    def __init__(self):
        osmium.SimpleHandler.__init__(self)
        self.xml_output = io.StringIO()

    def node(self, n):
        self.xml_output.write(f'<node id="{n.id}" lat="{n.location.lat}" lon="{n.location.lon}">\n')
        for k, v in n.tags:
            self.xml_output.write(f'  <tag k="{k}" v="{v}"/>\n')
        self.xml_output.write('</node>\n')

    def way(self, w):
        self.xml_output.write(f'<way id="{w.id}">\n')
        for n in w.nodes:
            self.xml_output.write(f'  <nd ref="{n.ref}"/>\n')
        for k, v in w.tags:
            self.xml_output.write(f'  <tag k="{k}" v="{v}"/>\n')
        self.xml_output.write('</way>\n')

    def relation(self, r):
        self.xml_output.write(f'<relation id="{r.id}">\n')
        for m in r.members:
            self.xml_output.write(f'  <member type="{m.type}" ref="{m.ref}" role="{m.role}"/>\n')
        for k, v in r.tags:
            self.xml_output.write(f'  <tag k="{k}" v="{v}"/>\n')
        self.xml_output.write('</relation>\n')

# Initialize the handler and apply it to the PBF file
handler = OSMHandler()
handler.apply_file("C:\\Users\\phams\\Downloads\\vietnam-latest.osm.pbf")

# Save the output to an XML file
with open('C:\\Users\\phams\\Downloads\\output.osm', 'w') as f:
    f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
    f.write('<osm version="0.6">\n')
    f.write(handler.xml_output.getvalue())
    f.write('</osm>\n')
