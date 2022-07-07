use derive_more::From;
use strong_xml::{XmlRead, XmlWrite};

use crate::document::{Paragraph, Table};
use crate::formatting::SectionProperty;
use crate::__xml_test_suites;

use super::SDT;

/// Document Body
///
/// This is the main document editing surface.
#[derive(Debug, Default, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
#[xml(tag = "w:body")]
pub struct Body<'a> {
    /// Specifies the contents of the body of the document.
    #[xml(child = "w:p", child = "w:tbl", child = "w:sectPr", child = "w:sdt")]
    pub content: Vec<BodyContent<'a>>,
}

impl<'a> Body<'a> {
    pub fn push<T: Into<BodyContent<'a>>>(&mut self, content: T) -> &mut Self {
        self.content.push(content.into());
        self
    }

    pub fn text(&self) -> String {
        let v: Vec<_> = self
            .content
            .iter()
            .filter_map(|content| match content {
                BodyContent::Paragraph(para) => Some(para.iter_text()),
                BodyContent::Table(_) => None,
                BodyContent::SectionProperty(_) => None,
                BodyContent::Sdt(_) => None,
            })
            .flatten()
            .map(|t| t.to_string())
            .collect();
        v.join("")
    }

    pub fn replace_text_simple<S>(&mut self, old: S, new: S)
    where
        S: AsRef<str>,
    {
        let dic = (old, new);
        let dic = vec![dic];
        let _d = self.replace_text(&dic);
    }

    pub fn replace_text<'b, T, S>(&mut self, dic: T) -> crate::DocxResult<()>
    where
        S: AsRef<str> + 'b,
        T: IntoIterator<Item = &'b (S, S)> + std::marker::Copy,
    {
        for content in self.content.iter_mut() {
            match content {
                BodyContent::Paragraph(p) => {
                    let _d = p.replace_text(dic)?;
                }
                BodyContent::Table(_) => {}
                BodyContent::SectionProperty(_) => {}
                BodyContent::Sdt(_) => {}
            }
        }
        Ok(())
    }

    // pub fn iter_text(&self) -> impl Iterator<Item = &Cow<'a, str>> {
    //     self.content
    //         .iter()
    //         .filter_map(|content| match content {
    //             BodyContent::Paragraph(para) => Some(para.iter_text()),
    //         })
    //         .flatten()
    // }

    // pub fn iter_text_mut(&mut self) -> impl Iterator<Item = &mut Cow<'a, str>> {
    //     self.content
    //         .iter_mut()
    //         .filter_map(|content| match content {
    //             BodyContent::Paragraph(para) => Some(para.iter_text_mut()),
    //         })
    //         .flatten()
    // }
}

/// A set of elements that can be contained in the body
#[derive(Debug, From, XmlRead, XmlWrite, Clone)]
#[cfg_attr(test, derive(PartialEq))]
pub enum BodyContent<'a> {
    #[xml(tag = "w:p")]
    Paragraph(Paragraph<'a>),
    #[xml(tag = "w:tbl")]
    Table(Table<'a>),
    #[xml(tag = "w:sdt")]
    Sdt(SDT<'a>),
    #[xml(tag = "w:sectPr")]
    SectionProperty(SectionProperty<'a>),
}

__xml_test_suites!(
    Body,
    Body::default(),
    r#"<w:body/>"#,
    Body {
        content: vec![Paragraph::default().into()]
    },
    r#"<w:body><w:p/></w:body>"#,
    Body {
        content: vec![Table::default().into()]
    },
    r#"<w:body><w:tbl><w:tblPr/></w:tbl></w:body>"#,
);
