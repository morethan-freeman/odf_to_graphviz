from zipfile import ZipFile
import re
import xml.etree.ElementTree as ET


# Encapsulates the list of namespaces use by the XML document being read. Can then be used
# to expand namespace names into their full values, allowing values to be searched for
# by their fully qualified names.
class namespace:

    def __init__(self, xml_root_data_as_string):
        # Read and store the contents of the namespace
        self._namespace_data = self.read_namespace(xml_root_data_as_string)

    # Read and return the contents if the namespace
    @staticmethod
    def read_namespace(xml_root_data_as_string):

        # For each namespace definition in the xml document
        namespace_temp = {}
        for match in re.finditer(" xmlns:", xml_root_data_as_string):
            # Pull off the name=value pair following the 'xmlns:'
            key_value_tri_str = xml_root_data_as_string[match.end():len(xml_root_data_as_string)].split('"', 2)
            # Split the name and value and add them into the dictionary
            namespace_temp[key_value_tri_str[0].split('=', 1)[0]] = key_value_tri_str[1]

        return namespace_temp

    # Expand the namespace name within a string and return the expanded result. This is useful for
    # searching XML documents for particular values via the etree module
    def expand_namespace(self, string_to_expand):
        [namespace_name, tag] = string_to_expand.split(':', 1)
        ns_name = ""
        if namespace_name in self._namespace_data:
            ns_name = "{" + self._namespace_data[namespace_name] + "}"
        return ns_name + tag


# Open the xml file that contains the main contents of the PowerPoint.


# Open the XML file that contains the main contents of the PowerPoint.
def read_the_extracted_xml_file(file_name):
    with open(file_name, encoding='utf-8') as data_file:
        xml_data = data_file.read()

    return xml_data

# Open the ODF file, decompress it and pull out the content of the relevant XML file which
# describes the content of the file.
def read_odf_file(file_name):
    with ZipFile(file_name, 'r') as zip:
        odf_data = zip.read("content.xml")

    # Convert the binary contents of the file into a string and return
    return odf_data.decode("utf-8")


# Parse from the root of the XML down to location where the network diagram is contained
def get_page_from_powerpoint_data(root):

    # Note: This function is customised for open document format (ODF) files saved from PowerPoint.
    # These files can also be generated in OpenOffice. To parse other ODF formats generated, by, for
    # example, Visio or OpenOffice Draw, minor changes to this function are required to navigate the
    # very slightly different XML tree structure in those formats. Note, it would also be easily possible
    # to scan a a list of all the pages allowing aggregation their contents as a single network diagram
    # if we amend this function to return a list of all pages instead of the first one found.

    # Extract the body from the document
    body = root.find(ns.expand_namespace('office:body'))

    # Extract the drawing (called drawing for open ofic draw, or presentation for powerpoints)
    drawing = body.find(ns.expand_namespace('office:presentation'))  # For powerpoints

    # Read the first page of the document
    page = drawing.find(ns.expand_namespace('draw:page'))

    return page


# Parses through the XML paragraphs that make up the contents of the text box within a shape, and sticthes them
# together
def get_shape_label(shape):
    label = ""

    # Read each line of the text comments from within the shape's text box
    paragraphs = shape.findall(ns.expand_namespace('text:p'))

    # For each paragraph
    for paragraph in paragraphs:

        # Get all of the snippets of text that exist within the paragraph
        snippets = paragraph.findall(ns.expand_namespace('text:span'))

        # Stitch all the snippets together into a single line
        for snippet in snippets:
            if snippet.text is not None:
                label += snippet.text

        # Add a newline at the end of each paragraph
        label += "\n"

    print(f'label: {label}')
    return label


# Get a list of all the shapes in the immediate level of the first page. Note we don't pull out the contents of groups,
# so shapes inside any groups will be ignored. Groups break the source-destination linkages of connectors so should be
# avoided when defining networks.
def get_shapes(page):

    # Read all the shapes (vms and ports and networks) off the page
    shapes = page.findall(ns.expand_namespace('draw:custom-shape'))

    # Read the relevant node data from each of the shapes
    scanned_shapes = []
    for shape in shapes:
        # Read the label by assembling together all the text contents of the box

        scanned_shapes.append({
            'id': shape.attrib[ns.expand_namespace('draw:id')],
            'shape': 'egg',
            'label': get_shape_label(shape)
        })

    return scanned_shapes


# Read all the connectors that join one shape to another. Connectors in PowerPoint have a source and destination
# object. We are concerned with connectors, not lines (lines have no source and destination)
def get_connectors(page):

    # Can be used to add a unique ID to the edge if it is missing
    edge_count = 0

    # Read all the unlabelled connectors off the page (joining vms to ports or networks to ports)
    edges = page.findall(ns.expand_namespace('draw:connector'))

    # Read the relevant node data from each of the shapes
    scanned_connectors = []
    for edge in edges:

        # Update the generator we use for our unique ID if one is not found in the connector
        edge_count += 1

        # Check if the unique ID is available
        try:
            connector_id = edge.attrib[ns.expand_namespace('draw:id')]

        # If it isn't, generate one ourselves
        except Exception:
            connector_id = f"con_id{edge_count}"

        # Read the label by assembling together all the text contents of the box
        try:
            scanned_connectors.append({
                'id': connector_id,
                'source': edge.attrib[ns.expand_namespace('draw:start-shape')],
                'destination': edge.attrib[ns.expand_namespace('draw:end-shape')],
                'label': ""
            })

        except Exception as e:
            raise Exception(f"Edge was: {edge}, Exception was {e}")

    return scanned_connectors


# Look up the source and destination of the connector and return it in the same object
def get_source_and_dest(connector, scanned_shapes):

    source = None
    dest = None
    source_id = connector['source']
    dest_id = connector['destination']
    for shape in scanned_shapes:

        # If the connector ID matches the source
        if shape['id'] == source_id:
            if source is not None:
                raise Exception("can't have more than one source for a connector")
            else:
                source = shape

        if shape['id'] == dest_id:
            if dest is not None:
                raise Exception("can't have more than one dest for a connector")
            else:
                dest = shape

    return source, dest


# Read the shape label, from which we can then extract its properties. FOr simplicity, a comma separated list of
# name=value pairs within the text label of a shape is used to define the poperties of the network aartefact
# represented by that shape.
def get_object_label(shape):
    return shape['label']


# Read a string defining a current object. Read each name=value pair into the label into a dictionary entry
def get_object_properties_from_label(label_string):
    label_string = label_string.replace('\\n', '')
    label_string = label_string.strip('"')
    try:
        props = dict((x.strip(), y.strip()) for x, y in (element.split('=') for element in label_string.split(',')))
    except Exception as e:
        raise Exception(f"\nLabel was in the wrong format. Require comma separated list of valid name-value pairs:\n"
                        f"Exact issue was {e}")

    return props


# Read the value of a specific parameter from the an object label
def get_object_parameter(shape, param_name):

    # Default is to return None if the param doesn't exist
    param_value = None

    # Read the object label
    label = get_object_label(shape)
    if label is not None:
        # If a label exists, parse its contents to create a dictionary of node properties
        object_properties = get_object_properties_from_label(label)

        if param_name in object_properties:
            param_value = object_properties[param_name]

    return param_value


# Check within the label of node or edge to see if a specified attribute is present with a specified value.
# Useful for searching for particular types of object (such as type=vm or type=port) within lists of objects
def matches_parameter_value_pair(item, param, value):
    match_found = False
    item_type = get_object_parameter(item, param)

    if item_type is not None:
        if item_type == value:
            match_found = True

    return match_found


# Look at the type of a shape and determine whether it is a port
def is_a_port(shape):
    return matches_parameter_value_pair(shape, "type", "port")


# Look at the type of a shape and determines whether it is a vm
def is_a_vm(shape):
    return matches_parameter_value_pair(shape, "type", "vm")


# Looks at the type of a shape and determines whether it is a network
def is_a_network(shape):
    return matches_parameter_value_pair(shape, "type", "net")


# Look at a list of network items, searching for one that has the described name-value pair.
# Returns its list index if found, otherwise returns None
def find_on_list(list_of_shapes, param_name, param_value):

    print(f"finding {param_name}={param_value} in {list_of_shapes}")
    index_of_found_item = None
    # Look through the list for a match, and if found get its index in the list
    for shape in list_of_shapes:
        if matches_parameter_value_pair(shape, param_name, param_value):
            index_of_found_item = list_of_shapes.index(shape)

    return index_of_found_item


# Look at a list of network items, searching for one that has specified ID. An ID is a unique string assigned
# to every drawn object by PowerPoint. IDs can be used in place of user-provided names
def find_on_list_by_id(list_of_shapes, param_value):

    index_of_found_item = None
    # Look through the list for a match, and if found get its index in the list
    for shape in list_of_shapes:
        if shape['id'] == param_value:
            index_of_found_item = list_of_shapes.index(shape)

    return index_of_found_item


# Checks whether a network is currently on a list of networks. If it is, does nothing,
# but if it isn't, adds it in
def include_in_network_list(list, network):

    print(f"adding network: {network}")
    # If it isn't on the list, add it.
    if find_on_list_by_id(list_of_shapes=list, param_value=network['id']) is None:
        print(f"adding shape {network['id']} to {list}")
        list.append({
            'id': network['id'],
            'shape': network['shape'],
            'label': network['label'],
            'ports': []
        })

    return list


# Assuming a specified network is already on the list, add a port to it.
def add_port_to_network_node(list, network, port):

    item_index = find_on_list_by_id(list_of_shapes=list, param_value=network['id'])

    # Check the thing we're adding to is in the list
    if item_index is None:
        raise Exception("Can't add port to a network that isn't on the list")

    # If it is then add it and return the updated list
    list[item_index]['ports'].append(port['id'])
    return list


# Get a list of network nodes, each node containing a list of ports connected to it. We will
# later (in another function) transform each of these into a separate graphviz edge. Typically we would
# have two ports connected to each network node, but the connection of more than two ports to a single
# network node is supported and will result in an A -- B -- C -- etc. line in the generated graphviz.
def get_networks(scanned_connectors, scanned_shapes):
    
    # Read through all the connectors bulding up a list of networks nodes and ports they connect to
    network_list = []
    for connector in scanned_connectors:
        print(f"considering connector {connector}")
        # If a connector joins a port to a network
        source, dest = get_source_and_dest(connector, scanned_shapes)
        print(f"source={source['id']}  destination={dest['id']}")
        if is_a_port(source) and is_a_network(dest):

            # Make a note of this network node if we haven't already
            network_list = include_in_network_list(list=network_list, network=dest)

            # Add the source to this network node's connectivity list
            network_list = add_port_to_network_node(list=network_list, network=dest, port=source)

        elif is_a_port(dest) and is_a_network(source):

            # Make a note of this network node if we haven't already
            network_list = include_in_network_list(list=network_list, network=source)

            # Add the source to this network node's connectivity list
            network_list = add_port_to_network_node(list=network_list, network=source, port=dest)

    return network_list


# Return a list of virtual machines and ports
def get_all_vm_and_port_nodes(scanned_shapes):

    list_of_non_network_nodes = []
    for shape in scanned_shapes:
        if is_a_port(shape) or is_a_vm(shape):
            list_of_non_network_nodes.append(shape)

    return list_of_non_network_nodes


# Return a lost of all the connectors in the network diagram that join vm's to ports. These connectors
# are used to devine which NICs (ports) belong to each virtual machine
def get_all_vm_to_port_edges(scanned_connectors, scanned_shapes):

    list_of_non_network_edges = []
    for connector in scanned_connectors:
        # Lookup the actual source and dest objects from their IDs held in the connector object
        source, dest = get_source_and_dest(connector, scanned_shapes)

        # If neither end of the connector is a network, add this to the list of non-network edges
        if not is_a_network(source) and not is_a_network(dest):
            list_of_non_network_edges.append(connector)

    return list_of_non_network_edges


# Powerpoint can be a bit untidy at times and it's possible that some extra paragraps, empty lines and preceding
# spaces are present in the files. We want to strip out this kind of stuff to get a clean list of comma separated
# name-value pairs that we can use to configure up our vm's and other objects
def tidy_text(text):
    lines = text.split('\n')
    cleaned_lines = []
    print(f"split lines = {lines}")
    for line in lines:
        line = line.replace('\\n', '')
        line = line.strip()

        # If there's anything left on the line by this stage add it to a new list
        if len(line) > 0:
            cleaned_lines.append(line)

    cleaned_line = '\\n'.join(cleaned_lines)
    return cleaned_line


# Generate and return a string that constitutes a single line of a graphviz (.dot) file defining a node.
# By calling this function repeatedly and appending the strings created to a lager body of text, we can
# build up the .dot file line by line
def add_graphviz_node(node):

    node_name = get_object_parameter(node, "name")
    if node_name is None:
        node_name = node['id']

    node_shape = node['shape']
    node_label = node['label']
    print(f'Node = "{node_name}"[shape = {node_shape}, label = "{tidy_text(node_label)}"]\n')
    return f'  "{node_name}"[shape = {node_shape}, label = "{tidy_text(node_label)}"]\n'


# Create and return a string containing the lines of a .dot file required to define all the
# vm nodes and all the port nodes of the network
def add_non_network_graphviz_nodes(non_network_nodes):
    result = ""

    for node in non_network_nodes:
        result += add_graphviz_node(node)

    return result


# Generate and return a string that constitutes a single line of a graphviz (.dot) file defining an unlabelled edge.
# By calling this function repeatedly and appending the strings created to a lager body of text, we can
# build up the .dot file line by line
def add_graphviz_edge(edge, scanned_shapes):

    source, dest = get_source_and_dest(edge, scanned_shapes)
    edge_source_name = get_object_parameter(source, "name")
    if edge_source_name is None:
        edge_source_name = source['id']

    edge_destination_name = get_object_parameter(dest, "name")
    if edge_destination_name is None:
        edge_destination_name = dest['id']

    return f'  "{edge_source_name}" -- "{edge_destination_name}"\n'


# Create and return a string containing the lines of a .dot file required to define all the
# unlabelled edges of the network. These edges join ports to the vm's that own them
def add_non_network_graphviz_edges(non_network_edges, scanned_shapes):
    result = ""

    for edge in non_network_edges:
        result += add_graphviz_edge(edge, scanned_shapes)

    return result


# Generate and return a string that constitutes a single line of a graphviz (.dot) file defining a labelled edge.
# By calling this function repeatedly and appending the strings created to a lager body of text, we can
# build up the .dot file line by line
def add_graphviz_network_edge(network, scanned_shapes):

    result = ""

    # For each port in the list of ports contained in that object
    for port_id in network['ports']:

        # Look up the port object from its ID in the port list and get its name
        port_index = find_on_list_by_id(scanned_shapes, port_id)
        if port_index is None:
            raise Exception(f"Could not find port % in system")
        port = scanned_shapes[port_index]
        port_name = get_object_parameter(port, "name")
        if port_name is None:
            port_name = port_id

        # Connect them all together:

        if result == "":
            # Start of the line, first port
            result += "  "
        else:
            # Continuing a line, appending ports
            result += " -- "

        # Add the name of the port
        result += f'"{port_name}"'

    # Add the network object label
    label = tidy_text(network['label'])
    result += f' [label="{label}"] \n'

    return result


# Create and return a string containing the lines of a .dot file required to define all the
# labelled edges of the network. These edges join ports to other ports and represent the networks.
# Each of these links is labelled with the name of the network. It is possible to have several
# edges with the same name. All edges with the same name will be considered logically joined.
def add_network_graphviz_edges(network_connectors):
    result = ""

    print(f"network_connectors {network_connectors}")
    for object in network_connectors:
        result += add_graphviz_network_edge(object, scanned_shapes)

    return result


# Return a string that constitutes the header of a graphviz .dot file
def add_dot_header():
    return "graph G {\n"


# Return a string that constitutes the footer of a graphviz .dot file
def add_dot_closer():
    return "}"


''' Below are the steps  required to generate the .dot file'''

# =====================================================================
print("STEP 0: READING THE OPEN DOCUMENT FORMAT SAVED POWERPOINT FILE")
# =====================================================================

#name_of_xml_content_file = 'SimpleNetwork2.xml'
name_of_odf_powerpoint_file = 'FourNodeExample.odp'

# Open the xml file that contains the main contents of the PowerPoint.
#data_as_string = read_the_extracted_xml_file(name_of_xml_content_file)  # test and dev purpposes

# Open the ODF file saved from PowerPoint.
data_as_string = read_odf_file(name_of_odf_powerpoint_file)
print(data_as_string)

# Read the namespace information out of XML the document. We'll use this when searching for tags later
ns = namespace(data_as_string)

# Read the root node of the XML document, from which we can then dive down to the first page on which we
# expect to find the network diagram of interest
root = ET.fromstring(data_as_string)

# =====================================================================
print("STEP 1: OPENING UP THE DOCUMENT AND GETTING TO THE RIGHT PAGE")
# =====================================================================

# Delve down into the document and pull out the page containing the network diagram
page = get_page_from_powerpoint_data(root)

# =====================================================================
print("STEP 2: SCANNING IN THE SHAPES AND CONNECTORS FROM THE DIAGRAM")
# =====================================================================

# Extract the relevant information from the nodes
scanned_shapes = get_shapes(page)
print(f"scanned_shapes = {scanned_shapes}")

# Extract the relevant information from the unlabelled connectors
scanned_connectors = get_connectors(page)
print(f"scanned_connectors = {scanned_connectors}")

# =====================================================================
print("STEP 3: COMPILING A LIST OF NETWORKS, VM'S AND PORTS")
# =====================================================================

# Extract the relevant information from the labelled connectors
network_connectors = get_networks(scanned_connectors, scanned_shapes)

# Extract the relevant information from the vm and port nodes
non_network_nodes = get_all_vm_and_port_nodes(scanned_shapes)

# Compiling a list of non-network connectors (port-to-vm)
non_network_edges = get_all_vm_to_port_edges(scanned_connectors, scanned_shapes)

# =========================================================================
print("STEP 4: BUILD UP THE TEXT OF THE GRAPHVIZ FILE FROM COLLECTED INFO")
# =========================================================================

# Add the header
result = ""
result += add_dot_header()

# Add the networks and ports
result += add_non_network_graphviz_nodes(non_network_nodes)

# Add the unlabelled edges that connect vm's to ports that they own
result += add_non_network_graphviz_edges(non_network_edges, scanned_shapes)

# Add the labelled edges that represent networks that connect ports to other ports
result += add_network_graphviz_edges(network_connectors)
result += add_dot_closer()

# Show the results file
print(f"\nAutomatically generated graphviz specification:\n\n{result}")
print("")

spec_filename = "specification.dot"
# =======================================================================
print(f"STEP 5: Write the graphviz data to file '{spec_filename}'")
# =======================================================================

# Save the file
with open(spec_filename, "wt") as text_file:
    text_file.write(result)


