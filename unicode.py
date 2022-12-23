import base64, re, quopri

class UnicodeReader:
    """
    encoded UTF-8
    """

    def __init__(self) -> None:
        pass

    def encoded(self, words: str):
        """Encode."""

        try:
            word_regex = r'=\?{1}(.+)\?{1}([B|Q])\?{1}(.+)\?{1}='

            charset, encoding, encoded_text = re.match(word_regex, 
                                                       words).groups()
            if encoding == 'B':
                byte_string = base64.b64decode(encoded_text)
            elif encoding == 'Q':
                byte_string = quopri.decodestring(encoded_text)
            return byte_string.decode(charset)
        except:
            return words