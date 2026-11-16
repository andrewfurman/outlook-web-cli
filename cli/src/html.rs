//! Email HTML to readable text, close to the Node.js `html-to-text` output the
//! earlier version used: 120-column wrapping, links as `text [href]`, and the href
//! hidden when it is the same as the link text.

use html2text::render::{TaggedLine, TextDecorator};

const MARK_START: char = '\u{27E6}';
const MARK_END: char = '\u{27E7}';

#[derive(Clone, Debug, Default)]
struct MailDecorator {
    link: Vec<String>,
}

impl TextDecorator for MailDecorator {
    type Annotation = ();

    fn decorate_link_start(&mut self, url: &str) -> (String, Self::Annotation) {
        self.link.push(url.to_string());
        (String::new(), ())
    }
    fn decorate_link_end(&mut self) -> String {
        match self.link.pop() {
            Some(url) if !url.is_empty() && !url.starts_with('#') => {
                format!("{MARK_START}{url}{MARK_END}")
            }
            _ => String::new(),
        }
    }
    fn decorate_em_start(&self) -> (String, Self::Annotation) {
        (String::new(), ())
    }
    fn decorate_em_end(&self) -> String {
        String::new()
    }
    fn decorate_strong_start(&self) -> (String, Self::Annotation) {
        (String::new(), ())
    }
    fn decorate_strong_end(&self) -> String {
        String::new()
    }
    fn decorate_strikeout_start(&self) -> (String, Self::Annotation) {
        (String::new(), ())
    }
    fn decorate_strikeout_end(&self) -> String {
        String::new()
    }
    fn decorate_code_start(&self) -> (String, Self::Annotation) {
        (String::new(), ())
    }
    fn decorate_code_end(&self) -> String {
        String::new()
    }
    fn decorate_preformat_first(&self) -> Self::Annotation {}
    fn decorate_preformat_cont(&self) -> Self::Annotation {}
    fn decorate_image(&mut self, _src: &str, title: &str) -> (String, Self::Annotation) {
        let t = title.trim();
        (
            if t.is_empty() {
                String::new()
            } else {
                format!("[{t}]")
            },
            (),
        )
    }
    fn header_prefix(&self, _level: usize) -> String {
        String::new()
    }
    fn quote_prefix(&self) -> String {
        "> ".into()
    }
    fn unordered_item_prefix(&self) -> String {
        " * ".into()
    }
    fn ordered_item_prefix(&self, i: i64) -> String {
        format!("{i}. ")
    }
    fn make_subblock_decorator(&self) -> Self {
        MailDecorator::default()
    }
    fn finalise(&mut self, _links: Vec<String>) -> Vec<TaggedLine<Self::Annotation>> {
        Vec::new()
    }
}

/// Plain text of an HTML email body.
pub fn html_to_text(html: &str) -> String {
    let rendered = html2text::config::with_decorator(MailDecorator::default())
        .string_from_read(html.as_bytes(), 120)
        .unwrap_or_else(|_| html.to_string());
    let mut out = String::with_capacity(rendered.len());
    let mut rest = rendered.as_str();
    while let Some(start) = rest.find(MARK_START) {
        let before = &rest[..start];
        let after = &rest[start + MARK_START.len_utf8()..];
        let Some(end) = after.find(MARK_END) else {
            out.push_str(before);
            rest = after;
            break;
        };
        let url = &after[..end];
        out.push_str(before);
        let same = out.trim_end().ends_with(url)
            || out.trim_end().ends_with(url.trim_start_matches("mailto:"));
        if !same {
            out.push_str(&format!(" [{url}]"));
        }
        rest = &after[end + MARK_END.len_utf8()..];
    }
    out.push_str(rest);
    let lines: Vec<&str> = out.lines().map(str::trim_end).collect();
    lines.join("\n").trim().to_string()
}

#[cfg(test)]
mod tests {
    use super::html_to_text;

    #[test]
    fn links_and_breaks() {
        let t = html_to_text(
            r#"<p>Line 1<br>Line 2</p><p>See <a href="https://contoso.com/a">the doc</a> or <a href="https://contoso.com">https://contoso.com</a>.</p>"#,
        );
        assert!(t.starts_with("Line 1\nLine 2"), "{t}");
        assert!(t.contains("the doc [https://contoso.com/a]"), "{t}");
        assert!(t.contains("or https://contoso.com."), "{t}");
    }
}
