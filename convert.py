import os
import zipfile
import json

import xmind
from xmind.core.topic import TopicElement


def read_gmind(path):
    g = zipfile.ZipFile(path)
    bytes = g.read("content.json")
    g.close()
    xmind = json.loads(bytes)
    # print(xmind["root"]["data"]["text"])
    # print(xmind["root"]["children"][0]["data"]["text"])
    return xmind


def generate_xmind(json_dict, root_topic: TopicElement):
    # json_dict is a list of children
    for c in json_dict:
        # create sub topic of root topic
        t = root_topic.addSubTopic()
        t.setTitle(c["data"]["text"])
        if "hyperlink" in c["data"]:
            t.setURLHyperlink(c["data"]["hyperlink"])
        # image as a link
        # if "image" in c["data"]:
            # t.setURLHyperlink(c["data"]["image"])
        if "note" in c["data"] is not None:
            t.setPlainNotes(c["data"]["note"])
        if "priority" in c["data"] is not None:
            t.addMarker("priority-"+str(c["data"]["priority"]))

        # recusively visit child
        if "children" in c:
            generate_xmind(c["children"], t)


def convert_gmind_to_xmind(path):
    filename = os.path.splitext(path)[0]
    xmindFilename = filename + ".xmind"
    print(path + " -> " + xmindFilename)
    json_dict = read_gmind(path)

    # load first sheet
    workbook = xmind.load(xmindFilename)
    sheet1 = workbook.getPrimarySheet()
    sheet1.setTitle("first sheet")

    # set the root title
    root_topic1 = sheet1.getRootTopic()
    root_topic1.setTitle(json_dict["root"]["data"]["text"])

    generate_xmind(json_dict["root"]["children"], root_topic1)
    xmind.save(workbook, path=xmindFilename)
    add_manifest(xmindFilename)


def add_manifest(path):
    xmindName = os.path.basename(path)
    nameWithoudExt = os.path.splitext(xmindName)[0]
    zipName = nameWithoudExt + ".zip"
    # xmind to zip
    os.rename(xmindName, zipName)
    # append file to zip
    zip = zipfile.ZipFile(zipName, "a")

    # create manifet.xml
    f = open("manifest.xml", "w")
    manifestXml = """
    <?xml version="1.0" encoding="UTF-8" standalone="no"?>
<manifest xmlns="urn:xmind:xmap:xmlns:manifest:1.0">
    <file-entry full-path="comments.xml" media-type=""/>
    <file-entry full-path="content.xml" media-type="text/xml"/>
    <file-entry full-path="markers/" media-type=""/>
    <file-entry full-path="markers/markerSheet.xml" media-type=""/>
    <file-entry full-path="META-INF/" media-type=""/>
    <file-entry full-path="META-INF/manifest.xml" media-type="text/xml"/>
    <file-entry full-path="meta.xml" media-type="text/xml"/>
    <file-entry full-path="styles.xml" media-type=""/>
</manifest>
    """
    f.write(manifestXml)
    f.close()

    # append to zip
    manifestFile = "./manifest.xml"
    zip.write(manifestFile, "META-INF//manifest.xml")
    zip.close()
    os.remove(f.name)
    # zip to xmind
    os.rename(zipName, xmindName)


convert_gmind_to_xmind("Test.gmind")

# files = os.listdir(".")
# for f in files:
#   if f.endswith(".gmind"):
#   print(f)
#   convert_gmind_to_xmind(f)
