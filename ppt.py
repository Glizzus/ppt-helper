from typing import Annotated, TypeVar
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml, CT_OfficeStyleSheet, CT_SRgbColor
from lxml import etree
from pydantic import BaseModel, conint, conlist


def rgb_tuple_to_hex(rgb_tuple: tuple[int, int, int]) -> str:
    return f"{rgb_tuple[0]:02X}{rgb_tuple[1]:02X}{rgb_tuple[2]:02X}"


RGBType = Annotated[
    conlist(item_type=conint(ge=0, le=255), min_length=3, max_length=3),
    "Array of three integers between 0 and 255 representing the RGB color"
]


class ColorChoice(BaseModel):
    color: tuple[int, int, int] = RGBType,
    rationale: str


class FontChoice(BaseModel):
    name: str
    size: int
    rationale: str


class ThemeFonts(BaseModel):
    header: FontChoice
    body: FontChoice


class ThemeColors(BaseModel):
    dark: ColorChoice
    light: ColorChoice
    accent1: ColorChoice
    accent2: ColorChoice
    accent3: ColorChoice
    accent4: ColorChoice
    accent5: ColorChoice
    accent6: ColorChoice
    hyperlink: ColorChoice
    followed_hyperlink: ColorChoice

    def to_pptx_format(self):
        return {
            "dk2": rgb_tuple_to_hex(self.dark.color),
            "lt2": rgb_tuple_to_hex(self.light.color),
            "accent1": rgb_tuple_to_hex(self.accent1.color),
            "accent2": rgb_tuple_to_hex(self.accent2.color),
            "accent3": rgb_tuple_to_hex(self.accent3.color),
            "accent4": rgb_tuple_to_hex(self.accent4.color),
            "accent5": rgb_tuple_to_hex(self.accent5.color),
            "accent6": rgb_tuple_to_hex(self.accent6.color),
            "hlink": rgb_tuple_to_hex(self.hyperlink.color),
            "folHlink": rgb_tuple_to_hex(self.followed_hyperlink.color)
        }


class Theme(BaseModel):
    colors: ThemeColors
    fonts: ThemeFonts


class Schema(BaseModel):
    background: ColorChoice
    theme: Theme


namespaces = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}


def set_fonts(theme_xml: CT_OfficeStyleSheet, header_font: str, body_font: str):
    font_root = theme_xml.find(".//a:fontScheme", namespaces=namespaces)
    if not isinstance(font_root, etree._Element):
        raise ValueError("Font scheme is not a _Element")

    def set_font(font_xpath: str, font_name: str):
        font_elems = font_root.xpath(font_xpath, namespaces=namespaces)
        if font_elems:
            font_elem = font_elems[0]
            latin = font_elem.xpath(".//a:latin", namespaces=namespaces)
            if latin:
                latin[0].set('typeface', font_name)

    set_font(".//a:majorFont", header_font)
    set_font(".//a:minorFont", body_font)


def get_color_scheme_elem(theme_xml: CT_OfficeStyleSheet, key: str):
    elem_list = theme_xml.xpath(f"//a:themeElements/a:clrScheme/a:{key}")
    if not elem_list:
        raise ValueError(f"{key} not found in theme")
    elem = elem_list[0]
    if not isinstance(elem, etree._Element):
        raise ValueError(f"{key} is not a _Element")
    return elem


def modify_secondary_color(theme_xml: CT_OfficeStyleSheet, key: str, color: str):
    elem = get_color_scheme_elem(theme_xml, key)
    srgb_clr = elem.find(".//a:srgbClr", namespaces=namespaces)
    if not isinstance(srgb_clr, CT_SRgbColor):
        raise ValueError("srgbClr is not a _Element")
    srgb_clr.set('val', color)


def create_ppt(response_schema: Schema, output_path: str):
    prs = Presentation()
    slide_master = prs.slide_master
    background_fill = slide_master.background.fill
    background_fill.solid()
    background_fill.fore_color.rgb = RGBColor(*response_schema.background.color)

    theme_part = slide_master.part.part_related_by(RELATIONSHIP_TYPE.THEME)
    theme_xml = parse_xml(theme_part.blob)
    if not isinstance(theme_xml, CT_OfficeStyleSheet):
        raise ValueError("Theme part is not a CT_OfficeStyleSheet")

    test_dict = response_schema.theme.colors.to_pptx_format()
    for key, color in test_dict.items():
        try:
            modify_secondary_color(theme_xml, key, color)
        except ValueError as e:
            print(e)

    set_fonts(theme_xml, response_schema.theme.fonts.header.name, response_schema.theme.fonts.body.name)
    modified_theme_blobs = etree.tostring(theme_xml, xml_declaration=True, encoding='UTF-8')
    theme_part._blob = modified_theme_blobs
    prs.save(output_path)


if __name__ == "__main__":
    schema = Schema.model_validate({
        "background": {
            "color": [255, 255, 255],
            "rationale": "White background"
        },
        "theme": {
            "colors": {
                "dark": {
                    "color": [0, 0, 0],
                    "rationale": "Black text"
                },
                "light": {
                    "color": [255, 255, 255],
                    "rationale": "White background"
                },
                "accent1": {
                    "color": [0, 0, 255],
                    "rationale": "Blue accent"
                },
                "accent2": {
                    "color": [255, 0, 0],
                    "rationale": "Red accent"
                },
                "accent3": {
                    "color": [0, 255, 0],
                    "rationale": "Green accent"
                },
                "accent4": {
                    "color": [255, 255, 0],
                    "rationale": "Yellow accent"
                },
                "accent5": {
                    "color": [0, 255, 255],
                    "rationale": "Cyan accent"
                },
                "accent6": {
                    "color": [255, 0, 255],
                    "rationale": "Magenta accent"
                },
                "hyperlink": {
                    "color": [0, 0, 255],
                    "rationale": "Blue hyperlink"
                },
                "followed_hyperlink": {
                    "color": [128, 0, 128],
                    "rationale": "Purple followed hyperlink"
                }
            },
            "fonts": {
                "header": {
                    "name": "Comic Sans MS",
                    "rationale": "Arial for headers"
                },
                "body": {
                    "name": "Comic Sans MS",
                    "rationale": "Arial for body text"
                },
            },
        }
    })

    create_ppt(schema, "C:/Users/Glizzus/Documents/dev_container_template.pptx")
