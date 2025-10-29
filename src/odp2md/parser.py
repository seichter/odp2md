import xml.dom.minidom as dom
import re, unicodedata
import os
import zipfile


from scope import Scope
from slide import Slide


class Parser:
    def __init__(self):
        self.slides = []
        self.currentSlide = None
        self.currentText = ""
        self.currentDepth = 0
        self.currentScope = Scope.NONE
        self.mediaDirectory = "media"

    def getTextFromNode(self, node):
        if node.nodeType == node.TEXT_NODE and len(str(node.data)) > 0:
            return node.data
        return None

    def hasAttributeWithValue(self, node, name, value):
        if node.attributes == None:
            return False
        for attribute_name, attribute_value in node.attributes.items():
            if attribute_name == name and attribute_value == value:
                return True
        return False

    def debugNode(self, node):
        # print('node ', node.tagName)
        pass

    def handlePage(self, node):
        # set new current slide
        self.currentSlide = Slide()
        self.currentSlide.name = node.attributes["draw:name"]
        # parse
        self.handleNode(node)
        # store
        self.slides.append(self.currentSlide)

    def slugify(self, value, allow_unicode=False):
        """
        Convert to ASCII if 'allow_unicode' is False. Convert spaces to hyphens.
        Remove characters that aren't alphanumerics, underscores, or hyphens.
        Convert to lowercase. Also strip leading and trailing whitespace.
        """
        value = str(value)
        if allow_unicode:
            value = unicodedata.normalize("NFKC", value)
        else:
            value = (
                unicodedata.normalize("NFKD", value)
                .encode("ascii", "ignore")
                .decode("ascii")
            )
        value = re.sub(r"[^\w\s-]", "", value.lower()).strip()
        return re.sub(r"[-\s]+", "-", value)

    def handleNode(self, node):
        if self.hasAttributeWithValue(node, "presentation:class", "title"):
            self.currentScope = Scope.TITLE
        elif self.hasAttributeWithValue(node, "presentation:class", "outline"):
            self.currentScope = Scope.OUTLINE

        if node.nodeName in ["draw:image", "draw:plugin"]:
            for k, v in node.attributes.items():
                if k == "xlink:href":
                    # get the extension
                    name, ext = os.path.splitext(v)
                    ext = ext.lower()
                    # now we create a new slug name for conversion
                    slug = self.slugify(self.currentSlide.title)
                    if len(slug) < 1:
                        slug = "slide-" + str(len(self.slides)) + "-image"
                    slug += "-" + str(len(self.currentSlide.media))
                    slug = (slug[:128]) if len(slug) > 128 else slug  # truncate

                    self.currentSlide.media.append(
                        (v, os.path.join(self.mediaDirectory, slug + ext))
                    )

        t = self.getTextFromNode(node)

        if t != None:
            if self.currentScope == Scope.OUTLINE:
                self.currentText += (" " * self.currentDepth) + "- " + t + "\n"
            elif self.currentScope == Scope.TITLE:
                self.currentSlide.title += t
            elif self.currentScope == Scope.IMAGES:
                pass
                # print('image title ',t)

        for c in node.childNodes:
            self.currentDepth += 1
            self.handleNode(c)
            self.currentDepth -= 1

    def handleDocument(self, dom):
        # we only need the pages
        pages = dom.getElementsByTagName("draw:page")
        # iterate pages
        for page in pages:
            self.currentDepth = 0
            self.currentSlide = Slide()
            self.handleNode(page)
            self.currentSlide.text = self.currentText
            self.slides.append(self.currentSlide)

            self.currentText = ""

    def open(self, fname, mediaDir="media", markdown=False, mediaExtraction=False):
        self.mediaDirectory = mediaDir

        # open odp file
        with zipfile.ZipFile(fname) as odp:
            info = odp.infolist()
            for i in info:
                if i.filename == "content.xml":
                    with odp.open("content.xml") as index:
                        doc = dom.parseString(index.read())
                        self.handleDocument(doc)

            # output markdown
            if markdown == True:
                for slide in self.slides:
                    print(slide)

            # generate files
            if mediaExtraction == True:
                for slide in self.slides:
                    for m, v in slide.media:
                        try:
                            odp.extract(m, ".")
                            if not os.path.exists(self.mediaDirectory):
                                os.makedirs(self.mediaDirectory)
                            os.rename(os.path.join(".", m), v)
                        except KeyError:
                            print("error finding media file ", m)
