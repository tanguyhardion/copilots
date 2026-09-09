import unittest

from copilots_app.services.powerpoint.parser import parse_dsl


class PowerPointParserTests(unittest.TestCase):
    def test_parse_slide_background_command(self):
        shapes = parse_dsl("slide background=a1\nrect left=0 top=0 width=10 height=10 color=a2")
        self.assertEqual(shapes[0]["type"], "slide")
        self.assertEqual(shapes[0]["background_color"], "theme_accent1")
        self.assertEqual(shapes[1]["type"], "rect")


if __name__ == "__main__":
    unittest.main()

