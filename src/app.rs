//! Application-Defined File Properties part
//!
//! The corresponding ZIP item is `/docProps/app.xml`.

use hard_xml::{XmlError, XmlRead, XmlReader, XmlResult, XmlWrite, XmlWriter};
use zip::read;
use std::{borrow::Cow, ptr::null};
use std::io::{Error, Write};

use crate::schema::{SCHEMAS_EXTENDED, SCHEMA_DOC_PROPS_V_TYPES, SCHEMA_XML};

#[derive(Debug, Clone)]
pub struct App<'a> {
    pub template: Option<Cow<'a, str>>,
    pub total_time: Option<Cow<'a, str>>,
    pub pages: Option<Cow<'a, str>>,
    pub words: Option<Cow<'a, str>>,
    pub characters: Option<Cow<'a, str>>,
    pub application: Option<Cow<'a, str>>,
    pub doc_security: Option<Cow<'a, str>>,
    pub lines: Option<Cow<'a, str>>,
    pub paragraphs: Option<Cow<'a, str>>,
    pub scale_crop: Option<Cow<'a, str>>,
    pub company: Option<Cow<'a, str>>,
    pub links_up_to_date: Option<Cow<'a, str>>,
    pub characters_with_spaces: Option<Cow<'a, str>>,
    pub shared_doc: Option<Cow<'a, str>>,
    pub hyperlinks_changed: Option<Cow<'a, str>>,
    pub app_version: Option<Cow<'a, str>>,
}

impl Default for App<'static> {
    fn default() -> App<'static> {
        App {
            template: Some("Normal.dotm".into()),
            total_time: Some("1".into()),
            pages: Some("1".into()),
            words: Some("0".into()),
            characters: Some("0".into()),
            application: Some("docx-rs".into()),
            doc_security: Some("0".into()),
            lines: Some("0".into()),
            paragraphs: Some("1".into()),
            scale_crop: Some("false".into()),
            company: Some("MS".into()),
            links_up_to_date: Some("false".into()),
            characters_with_spaces: Some("25".into()),
            shared_doc: Some("false".into()),
            hyperlinks_changed: Some("false".into()),
            app_version: Some("12.0000".into()),
        }
    }
}

impl<'a> XmlWrite for App<'a> {
    fn to_writer<W: Write>(&self, writer: &mut XmlWriter<W>) -> XmlResult<()> {
        let App {
            template,
            total_time,
            pages,
            words,
            characters,
            application,
            doc_security,
            lines,
            paragraphs,
            scale_crop,
            company,
            links_up_to_date,
            characters_with_spaces,
            shared_doc,
            hyperlinks_changed,
            app_version,
        } = self;

        log::debug!("[App] Started writing.");

        let _ = write!(writer.inner, "{}", SCHEMA_XML);

        writer.write_element_start("Properties")?;

        writer.write_attribute("xmlns", SCHEMAS_EXTENDED)?;
        writer.write_attribute("xmlns:vt", SCHEMA_DOC_PROPS_V_TYPES)?;

        if template.is_none()
            && total_time.is_none()
            && pages.is_none()
            && words.is_none()
            && characters.is_none()
            && application.is_none()
            && doc_security.is_none()
            && lines.is_none()
            && paragraphs.is_none()
            && scale_crop.is_none()
            && company.is_none()
            && links_up_to_date.is_none()
            && characters_with_spaces.is_none()
            && shared_doc.is_none()
            && hyperlinks_changed.is_none()
            && app_version.is_none()
        {
            writer.write_element_end_empty()?;
        } else {
            writer.write_element_end_open()?;
            if let Some(val) = template {
                writer.write_flatten_text("Template", val, false)?;
            }
            if let Some(val) = total_time {
                writer.write_flatten_text("TotalTime", val, false)?;
            }
            if let Some(val) = pages {
                writer.write_flatten_text("Pages", val, false)?;
            }
            if let Some(val) = words {
                writer.write_flatten_text("Words", val, false)?;
            }
            if let Some(val) = characters {
                writer.write_flatten_text("Characters", val, false)?;
            }
            if let Some(val) = application {
                writer.write_flatten_text("Application", val, false)?;
            }
            if let Some(val) = doc_security {
                writer.write_flatten_text("DocSecurity", val, false)?;
            }
            if let Some(val) = lines {
                writer.write_flatten_text("Lines", val, false)?;
            }
            if let Some(val) = paragraphs {
                writer.write_flatten_text("Paragraphs", val, false)?;
            }
            if let Some(val) = scale_crop {
                writer.write_flatten_text("ScaleCrop", val, false)?;
            }
            if let Some(val) = company {
                writer.write_flatten_text("Company", val, false)?;
            }
            if let Some(val) = links_up_to_date {
                writer.write_flatten_text("LinksUpToDate", val, false)?;
            }
            if let Some(val) = characters_with_spaces {
                writer.write_flatten_text("CharactersWithSpaces", val, false)?;
            }
            if let Some(val) = shared_doc {
                writer.write_flatten_text("SharedDoc", val, false)?;
            }
            if let Some(val) = hyperlinks_changed {
                writer.write_flatten_text("HyperlinksChanged", val, false)?;
            }
            if let Some(val) = app_version {
                writer.write_flatten_text("AppVersion", val, false)?;
            }
            writer.write_element_end_close("Properties")?;
        }

        log::debug!("[App] Finished writing.");

        Ok(())
    }
}

impl<'a, 'b> XmlRead<'b> for App<'b> {
    fn from_reader(reader: &mut XmlReader<'b>) -> XmlResult<Self> {
        reader.read_till_element_start("")?;
        reader.read_till_element_start("Properties");
        Err(XmlError::IO(Error::new(std::io::ErrorKind::AddrInUse, "what")))
    }

    fn from_str(text: &'b str) -> XmlResult<Self> {
        let mut reader = XmlReader::new(text);
        Self::from_reader(&mut reader)
    }
}

#[cfg(test)]
mod test {
    use hard_xml::XmlRead;

    use super::App;


#[test]
fn read_old_app() {
let old_app_version = "<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?><ap:Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\" xmlns:ap=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><ap:Template>Normal.dotm</ap:Template><ap:Application>Microsoft Word for the web</ap:Application><ap:DocSecurity>0</ap:DocSecurity><ap:ScaleCrop>false</ap:ScaleCrop><ap:Company /><ap:SharedDoc>false</ap:SharedDoc><ap:HyperlinksChanged>false</ap:HyperlinksChanged><ap:AppVersion>16.0000</ap:AppVersion><ap:LinksUpToDate>false</ap:LinksUpToDate></ap:Properties>";
App::from_str(&old_app_version).expect("should exist");
}
}
