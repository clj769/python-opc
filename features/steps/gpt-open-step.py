from behave import given, when, then


from opc import OpcPackage
from opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT


# Global package object for steps to share
package = None

@given('a python-opc working environment')
def step_given_python_opc_environment(context):
    # This assumes that python-opc is correctly installed and ready to use.
    pass

@when('I open a PowerPoint file')
def step_when_open_ppt_file(context):
    global package
    # Replace 'filename.pptx' with your actual PowerPoint file path
    package = OpcPackage.open('tests/test_files/test.pptx')

@then('the expected package rels are loaded')
def step_then_package_rels_loaded(context):
    global package
    # You need to replace 'expected_rels' with your actual expected rels
    expected_rels = (
        ('rId1', RT.OFFICE_DOCUMENT,     False, '/ppt/presentation.xml'),
        ('rId2', RT.THUMBNAIL,           False, '/docProps/thumbnail.jpeg'),
        ('rId3', RT.CORE_PROPERTIES,     False, '/docProps/core.xml'),
        ('rId4', RT.EXTENDED_PROPERTIES, False, '/docProps/app.xml'),
    )

    assert package.rels == expected_rels, 'Package rels not as expected.'

@then('the expected parts are loaded')
def step_then_parts_are_loaded(context):
    global package
    # You need to replace 'expected_parts' with your actual expected parts
    expected_parts = {
        '/docProps/app.xml': (
            CT.OFC_EXTENDED_PROPERTIES, 'e5a7552c35180b9796f2132d39bc0d208cf'
            '8761f', []
        ),
        '/docProps/core.xml': (
            CT.OPC_CORE_PROPERTIES, '08c8ff0912231db740fa1277d8fa4ef175a306e'
            '4', []
        ),
        '/docProps/thumbnail.jpeg': (
            CT.JPEG, '8a93420017d57f9c69f802639ee9791579b21af5', []
        ),
        '/ppt/presentation.xml': (
            CT.PML_PRESENTATION_MAIN,
            'efa7bee0ac72464903a67a6744c1169035d52a54',
            [
                ('rId1', RT.SLIDE_MASTER, False,
                 '/ppt/slideMasters/slideMaster1.xml'),
                ('rId2', RT.SLIDE, False, '/ppt/slides/slide1.xml'),
                ('rId3', RT.PRINTER_SETTINGS, False,
                 '/ppt/printerSettings/printerSettings1.bin'),
                ('rId4', RT.PRES_PROPS, False, '/ppt/presProps.xml'),
                ('rId5', RT.VIEW_PROPS, False, '/ppt/viewProps.xml'),
                ('rId6', RT.THEME, False, '/ppt/theme/theme1.xml'),
                ('rId7', RT.TABLE_STYLES, False, '/ppt/tableStyles.xml'),
            ]
        ),
        '/ppt/printerSettings/printerSettings1.bin': (
            CT.PML_PRINTER_SETTINGS, 'b0feb4cc107c9b2d135b1940560cf8f045ffb7'
            '46', []
        ),
        '/ppt/presProps.xml': (
            CT.PML_PRES_PROPS, '7d4981fd742429e6b8cc99089575ac0ee7db5194', []
        ),
        '/ppt/viewProps.xml': (
            CT.PML_VIEW_PROPS, '172a42a6be09d04eab61ae3d49eff5580a4be451', []
        ),
        '/ppt/theme/theme1.xml': (
            CT.OFC_THEME, '9f362326d8dc050ab6eef7f17335094bd06da47e', []
        ),
        '/ppt/tableStyles.xml': (
            CT.PML_TABLE_STYLES, '49bfd13ed02199b004bf0a019a596f127758d926',
            []
        ),
        '/ppt/slideMasters/slideMaster1.xml': (
            CT.PML_SLIDE_MASTER, 'be6fe53e199ef10259227a447e4ac9530803ecce',
            [
                ('rId1', RT.SLIDE_LAYOUT, False,
                 '/ppt/slideLayouts/slideLayout1.xml'),
                ('rId2', RT.SLIDE_LAYOUT, False,
                 '/ppt/slideLayouts/slideLayout2.xml'),
                ('rId3', RT.SLIDE_LAYOUT, False,
                 '/ppt/slideLayouts/slideLayout3.xml'),
                ('rId4', RT.THEME, False, '/ppt/theme/theme1.xml'),
            ],
        ),
        '/ppt/slideLayouts/slideLayout1.xml': (
            CT.PML_SLIDE_LAYOUT, 'bcbeb908e22346fecda6be389759ca9ed068693c',
            [
                ('rId1', RT.SLIDE_MASTER, False,
                 '/ppt/slideMasters/slideMaster1.xml'),
            ],
        ),
        '/ppt/slideLayouts/slideLayout2.xml': (
            CT.PML_SLIDE_LAYOUT, '316d0fb0ce4c3560fa2ed4edc3becf2c4ce84b6b',
            [
                ('rId1', RT.SLIDE_MASTER, False,
                 '/ppt/slideMasters/slideMaster1.xml'),
            ],
        ),
        '/ppt/slideLayouts/slideLayout3.xml': (
            CT.PML_SLIDE_LAYOUT, '5b704e54c995b7d1bd7d24ef996a573676cc15ca',
            [
                ('rId1', RT.SLIDE_MASTER, False,
                 '/ppt/slideMasters/slideMaster1.xml'),
            ],
        ),
        '/ppt/slides/slide1.xml': (
            CT.PML_SLIDE, '1841b18f1191629c70b7176d8e210fa2ef079d85',
            [
                ('rId1', RT.SLIDE_LAYOUT, False,
                 '/ppt/slideLayouts/slideLayout1.xml'),
                ('rId2', RT.HYPERLINK, True,
                 'https://github.com/scanny/python-pptx'),
            ]
        ),
    }

    assert package.parts == expected_parts, 'Package parts not as expected.'
