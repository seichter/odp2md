import textwrap


class Slide:
    def __init__(self):
        self.title = ""
        self.text = ""
        self.notes = ""
        self.media = []

    def generateMarkdown(self, blockToHTML=True):
        # fix identation
        self.text = textwrap.dedent(self.text)
        out = "## {0}\n\n{1}\n".format(self.title, self.text)
        for m, v in self.media:
            # maybe let everything else fail?
            isVideo = any(x in v for x in [".mp4", ".mkv"])

            if blockToHTML and isVideo:
                # since LaTeX extensions for video are deprecated
                out += "`![]({0})`{{=html}}\n".format(v)
            else:
                out += "![]({0})\n".format(v)
        return out

    # override string representation
    def __str__(self):
        return self.generateMarkdown()
