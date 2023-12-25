use std::borrow::Cow;
use hard_xml::{XmlRead, XmlWrite};

use crate::{__setter, __xml_test_suites, formatting::BorderStyle};

#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:bottom")]
pub struct BottomBorder<'a> {
    #[xml(attr = "w:val")]
    pub style: super::BorderStyle,
    #[xml(attr = "w:color")]
    pub color: Option<Cow<'a, str>>,
    #[xml(attr = "w:themeColor")]
    pub theme_color: Option<crate::formatting::ThemeColor>,
    #[xml(attr = "w:themeTint")]
    pub theme_tint: Option<Cow<'a, str>>,
    #[xml(attr = "w:themeShade")]
    pub theme_shade: Option<Cow<'a, str>>,
    #[xml(attr = "w:sz")]
    pub size: Option<isize>, // Measurement in Eighths of a Point
    #[xml(attr = "w:space")]
    pub space: Option<isize>,
    #[xml(attr = "w:shadow")]
    pub shadow: Option<bool>,
    #[xml(attr = "w:frame")]
    pub frame: Option<bool>,
}

impl<'a> BottomBorder<'a> {
    __setter!(color: Option<Cow<'a, str>>);
    __setter!(shadow: Option<bool>);
    __setter!(space: Option<isize>);
    __setter!(size: Option<isize>);
    __setter!(style: BorderStyle);
}

__xml_test_suites!(
    BottomBorder,
    BottomBorder::default(),
    r#"<w:bottom w:val="none"/>"#,
    BottomBorder::default().color("000000"),
    r#"<w:bottom w:val="none" w:color="000000"/>"#,
    BottomBorder::default().shadow(false),
    r#"<w:bottom w:val="none" w:shadow="false"/>"#,
    BottomBorder::default().space(40isize),
    r#"<w:bottom w:val="none" w:space="40"/>"#,
    BottomBorder::default().size(20isize),
    r#"<w:bottom w:val="none" w:sz="20"/>"#,
    BottomBorder::default().style(BorderStyle::Dotted),
    r#"<w:bottom w:val="dotted"/>"#,
);
