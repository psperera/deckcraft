//! U30: native charts drawn with PowerPoint's plot geometry. A manual inner plot layout puts the
//! plot exactly where the file says, fixed axis ends are kept exactly, and deleted axes draw
//! nothing, also after the chart is edited. U31: a bar chart's category labels fit beside its
//! plot, and points without a value are gaps rather than zeros. Each chart is XML written here,
//! imported and rendered; no Office file is used.

use std::io::{Cursor, Read, Write};
use std::sync::Arc;

use deckcraft_engine::render::{Image, RenderOpts, render_slide};
use deckcraft_geom::Xfrm;
use deckcraft_model::chart::{Chart, ChartType, PlotLayout};
use deckcraft_model::{LayoutType, Presentation, Shape, ShapeId, ShapeKind, defaults};

const NS: &str =
    r#"xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;

/// The chart frame on the 960 × 540 slide: x, y, width, height.
const FRAME: (f64, f64, f64, f64) = (40.0, 30.0, 800.0, 450.0);

/// The inner plot at 10 % / 20 % of the frame, 80 % wide and 60 % high: (120, 120) to (760, 390)
/// on the slide.
const LAYOUT: &str = r#"<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.1"/><c:y val="0.2"/><c:w val="0.8"/><c:h val="0.6"/></c:manualLayout></c:layout>"#;
const PLOT: (f64, f64, f64, f64) = (120.0, 120.0, 760.0, 390.0);

const RED: [u8; 3] = [255, 0, 0];

/// The axis settings that vary between the tests.
#[derive(Clone, Copy)]
struct Axes<'a> {
    /// `c:scaling` children of the value axis other than the orientation.
    scaling: &'a str,
    reversed: bool,
    /// The value axis' `c:delete` and `c:tickLblPos`.
    deleted: bool,
    tick_labels: &'a str,
    major: Option<f64>,
    /// The category axis' `c:delete` and `c:tickLblPos`.
    cat_deleted: bool,
    cat_labels: &'a str,
}

/// Both axes shown, the value axis fixed at −10 to 44.
const SHOWN: Axes<'static> =
    Axes { scaling: FIXED, reversed: false, deleted: false, tick_labels: "nextTo", major: None, cat_deleted: false, cat_labels: "nextTo" };

/// A clustered column chart of one red series, `values` in categories Q1, Q2…, with a gap
/// width of 60 %, 10 pt text and no title or legend.
fn column_chart(values: &[f64], layout: &str, ax: Axes<'_>) -> String {
    clustered("col", values, layout, ax)
}

/// [`column_chart`] with its bars running across (`c:barDir val="bar"`): the category axis on
/// the left, first category at the bottom, and the value axis along the bottom.
fn bar_chart(values: &[f64], layout: &str, ax: Axes<'_>) -> String {
    clustered("bar", values, layout, ax)
}

fn clustered(dir: &str, values: &[f64], layout: &str, ax: Axes<'_>) -> String {
    let (cat_pos, val_pos) = if dir == "bar" { ("l", "b") } else { ("b", "l") };
    let n = values.len();
    let pts: String = values.iter().enumerate().map(|(i, v)| format!(r#"<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>"#)).collect();
    let cats: String = (0..n).map(|i| format!(r#"<c:pt idx="{i}"><c:v>Q{}</c:v></c:pt>"#, i + 1)).collect();
    let orientation = if ax.reversed { "maxMin" } else { "minMax" };
    let Axes { scaling, tick_labels, cat_labels, .. } = ax;
    let (delete, cat_delete) = (u8::from(ax.deleted), u8::from(ax.cat_deleted));
    let major = ax.major.map(|m| format!(r#"<c:majorUnit val="{m}"/>"#)).unwrap_or_default();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace {NS}><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>{layout}<c:barChart><c:barDir val="{dir}"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/><c:cat><c:strRef><c:strCache><c:ptCount val="{n}"/>{cats}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{n}"/>{pts}</c:numCache></c:numRef></c:val></c:ser><c:gapWidth val="60"/><c:axId val="11"/><c:axId val="12"/></c:barChart><c:catAx><c:axId val="11"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="{cat_delete}"/><c:axPos val="{cat_pos}"/><c:tickLblPos val="{cat_labels}"/><c:crossAx val="12"/><c:crosses val="autoZero"/></c:catAx><c:valAx><c:axId val="12"/><c:scaling><c:orientation val="{orientation}"/>{scaling}</c:scaling><c:delete val="{delete}"/><c:axPos val="{val_pos}"/><c:numFmt formatCode="General" sourceLinked="0"/><c:tickLblPos val="{tick_labels}"/><c:crossAx val="11"/><c:crosses val="autoZero"/><c:crossBetween val="between"/>{major}</c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:chartSpace>"#
    )
}

/// A one-slide deck whose only shape is a chart in `FRAME` with `xml` as its chart part: a deck
/// is exported with a stand-in chart, its chart part replaced, and the package imported again.
fn deck(xml: &str) -> Presentation {
    let mut p = defaults::new_presentation(None);
    p.slides.clear();
    let layout = defaults::layout_of_kind(&p, LayoutType::Blank).unwrap_or_default();
    let mut slide = defaults::new_slide(&mut p, layout).expect("slide");
    let (x, y, w, h) = FRAME;
    let mut shape = Shape { id: ShapeId(p.alloc_id().expect("id")), name: "Chart".into(), xfrm: Some(Xfrm::new(x, y, w, h)), ..Default::default() };
    shape.kind = ShapeKind::Chart(Box::new(Chart::sample(ChartType::Column)));
    slide.shapes.push(shape);
    p.slides.push(Arc::new(slide));
    let bytes = deckcraft_pptx::export(&p).expect("export");
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).expect("zip");
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = zip::write::SimpleFileOptions::default();
    let mut replaced = 0;
    for i in 0..zip.len() {
        let mut entry = zip.by_index(i).expect("entry");
        let name = entry.name().to_string();
        let mut data = vec![];
        entry.read_to_end(&mut data).expect("read");
        if name.starts_with("ppt/charts/chart") && name.ends_with(".xml") {
            data = xml.as_bytes().to_vec();
            replaced += 1;
        }
        out.start_file(&name, opts).expect("start");
        out.write_all(&data).expect("write");
    }
    assert_eq!(replaced, 1, "one chart part");
    deckcraft_pptx::import(&out.finish().expect("finish").into_inner()).expect("import")
}

fn chart_of(p: &Presentation) -> &Chart {
    p.slides[0].shapes.iter().find_map(|s| if let ShapeKind::Chart(c) = &s.kind { Some(&**c) } else { None }).expect("a chart")
}

fn image(p: &Presentation) -> Image {
    let img = render_slide(p, 0, &RenderOpts { scale: 1.0, ..Default::default() });
    assert_eq!((img.width, img.height), (960, 540));
    img
}

fn rgb(img: &Image, x: u32, y: u32) -> [u8; 3] {
    let p = img.pixel(x, y);
    [p[0], p[1], p[2]]
}

/// The bars: each run of exactly red columns, as (left, top, right, bottom), right and bottom
/// exclusive, so a bar from x = 150 to 250 reads 150 to 250.
fn bars(img: &Image) -> Vec<(u32, u32, u32, u32)> {
    let red_x = |x: u32| (0..img.height).any(|y| rgb(img, x, y) == RED);
    let mut out = vec![];
    let mut x = 0;
    while x < img.width {
        if !red_x(x) {
            x += 1;
            continue;
        }
        let x0 = x;
        while x < img.width && red_x(x) {
            x += 1;
        }
        let mid = (x0 + x) / 2;
        let ys: Vec<u32> = (0..img.height).filter(|y| rgb(img, mid, *y) == RED).collect();
        out.push((x0, ys[0], x, ys[ys.len() - 1] + 1));
    }
    out
}

/// Whether anything but the white slide shows in the box (inclusive).
fn ink(img: &Image, (x0, y0, x1, y1): (u32, u32, u32, u32)) -> bool {
    (y0..=y1).any(|y| (x0..=x1).any(|x| rgb(img, x, y).iter().any(|c| *c < 0xF0)))
}

/// Dark text ink (the chart's 595959 default) in the box (inclusive).
fn text(img: &Image, (x0, y0, x1, y1): (u32, u32, u32, u32)) -> bool {
    (y0..=y1).any(|y| (x0..=x1).any(|x| rgb(img, x, y).iter().all(|c| *c < 0x90)))
}

fn near(got: u32, want: f64) -> bool {
    (got as f64 - want).abs() <= 1.0
}

/// Where PowerPoint puts category `i`'s column of `n`: slots of plot width ÷ n, each column the
/// slot ÷ (1 + gap width), centred in its slot. Returns (left, right).
fn column_x(i: usize, n: usize, gap_width: f64) -> (f64, f64) {
    let (x0, _, x1, _) = PLOT;
    let slot = (x1 - x0) / n as f64;
    let bar = slot / (1.0 + gap_width / 100.0);
    let mid = x0 + slot * (i as f64 + 0.5);
    (mid - bar / 2.0, mid + bar / 2.0)
}

/// Where value `v` sits on an axis from `min` to `max` over the plot's height: the minimum at the
/// bottom, or at the top when the axis is reversed.
fn value_y(v: f64, (min, max): (f64, f64), reversed: bool) -> f64 {
    let (_, y0, _, y1) = PLOT;
    let t = (v - min) / (max - min);
    if reversed { y0 + t * (y1 - y0) } else { y1 - t * (y1 - y0) }
}

const VALUES: [f64; 4] = [12.0, 30.0, -6.0, 44.0];
const RANGE: (f64, f64) = (-10.0, 44.0);
const FIXED: &str = r#"<c:max val="44"/><c:min val="-10"/>"#;

/// Checks every column of `VALUES` against PowerPoint's arithmetic on the manual plot.
fn assert_columns_placed(img: &Image, reversed: bool) {
    let got = bars(img);
    assert_eq!(got.len(), VALUES.len(), "{got:?}");
    for (i, (&(l, t, r, b), &v)) in got.iter().zip(&VALUES).enumerate() {
        let (want_l, want_r) = column_x(i, VALUES.len(), 60.0);
        assert!(near(l, want_l) && near(r, want_r), "column {i}: x {l}..{r}, want {want_l}..{want_r}");
        // Each column runs from zero (where the category axis crosses) to its value.
        let (a, z) = (value_y(0.0, RANGE, reversed), value_y(v, RANGE, reversed));
        let (want_t, want_b) = (a.min(z), a.max(z));
        assert!(near(t, want_t) && near(b, want_b), "column {i} ({v}): y {t}..{b}, want {want_t}..{want_b}");
    }
}

/// The strategy kit's decorated charts and the add-in's charts: a manual inner plot, the value
/// axis fixed and deleted. Column centres and tops land where shapes drawn over the chart
/// expect them.
#[test]
fn a_manual_inner_layout_places_columns_by_powerpoints_arithmetic() {
    let ax = Axes { deleted: true, ..SHOWN };
    let p = deck(&column_chart(&VALUES, LAYOUT, ax));
    let c = chart_of(&p);
    assert_eq!(c.plot_layout, Some(PlotLayout { x: 0.1, y: 0.2, w: 0.8, h: 0.6 }));
    assert_eq!((c.y_min, c.y_max, c.y_hidden), (Some(-10.0), Some(44.0), true));
    let img = image(&p);
    assert_columns_placed(&img, false);
    // Spot values, 5 px to a unit: the second column (30) is centred at 360 and tops out at
    // 390 − 40 × 5 = 190; the last (44, the fixed maximum) reaches the plot's top edge exactly.
    let got = bars(&img);
    assert!(near((got[1].0 + got[1].2) / 2, 360.0) && near(got[1].1, 190.0), "{:?}", got[1]);
    assert!(near(got[3].1, PLOT.1), "{:?}", got[3]);
    // The category labels sit under the manual plot, centred under their columns.
    assert!(text(&img, (180, 393, 220, 410)), "Q1 under its column");

    // Without the manual layout the same chart is laid out automatically, elsewhere.
    let auto = image(&deck(&column_chart(&VALUES, "<c:layout/>", ax)));
    assert_ne!(bars(&auto), got, "the automatic layout differs");
}

/// U29's reversed value axis on the same manual plot: the columns hang from zero near the top,
/// in the same slots.
#[test]
fn a_manual_inner_layout_works_with_a_reversed_value_axis() {
    let ax = Axes { reversed: true, deleted: true, ..SHOWN };
    let p = deck(&column_chart(&VALUES, LAYOUT, ax));
    assert!(chart_of(&p).y_reversed);
    let img = image(&p);
    assert_columns_placed(&img, true);
    // The fixed minimum is the top edge: the last column (44) hangs to the bottom edge.
    let got = bars(&img);
    assert!(near(got[3].3, PLOT.3), "{:?}", got[3]);
    // With a reversed axis the category labels go above the plot.
    assert!(text(&img, (180, 100, 220, 117)), "Q1 above its column");
}

/// `c:min`/`c:max` are the axis' ends: the tallest column reaches the plot's top edge, and the
/// ticks stop at the last major unit inside the axis rather than labelling a rounded-up maximum.
#[test]
fn fixed_axis_ends_are_drawn_exactly() {
    let values = [100.0, 288.0, 150.0, 50.0];
    let ax = Axes { scaling: r#"<c:max val="288"/><c:min val="0"/>"#, major: Some(50.0), ..SHOWN };
    let p = deck(&column_chart(&values, LAYOUT, ax));
    let img = image(&p);
    let got = bars(&img);
    assert_eq!(got.len(), 4, "{got:?}");
    for (i, (&(_, t, _, b), &v)) in got.iter().zip(&values).enumerate() {
        assert!(near(t, value_y(v, (0.0, 288.0), false)) && near(b, PLOT.3), "column {i} ({v}): y {t}..{b}");
    }
    assert!(near(got[1].1, PLOT.1), "288 reaches the top edge: {:?}", got[1]);
    // Tick labels left of the plot at 0, 50 … 250; none at the top edge, where a maximum
    // rounded up to 300 would put one.
    let gutter = |y: f64| (40, (y - 7.0) as u32, PLOT.0 as u32 - 2, (y + 7.0) as u32);
    for v in [0.0, 50.0, 100.0, 150.0, 200.0, 250.0] {
        assert!(text(&img, gutter(value_y(v, (0.0, 288.0), false))), "a label at {v}");
    }
    assert!(!text(&img, gutter(PLOT.1)), "no label at the top edge (288 is not a major unit)");
}

/// A deleted value axis, and one whose tick labels are off (`c:tickLblPos none`), draws no line
/// and no labels; the plot keeps its place.
#[test]
fn a_deleted_value_axis_draws_nothing() {
    let shown = SHOWN;
    // Left of the plot, from above its top to below its bottom (the category labels are under
    // the columns, further right).
    let strip = (FRAME.0 as u32, PLOT.1 as u32 - 10, PLOT.0 as u32 - 1, PLOT.3 as u32 + 10);
    let img = image(&deck(&column_chart(&VALUES, LAYOUT, shown)));
    assert!(text(&img, strip), "a shown axis has its labels there");
    assert_columns_placed(&img, false);
    for hidden in [Axes { deleted: true, ..shown }, Axes { tick_labels: "none", ..shown }] {
        let p = deck(&column_chart(&VALUES, LAYOUT, hidden));
        assert!(chart_of(&p).y_hidden);
        let img = image(&p);
        assert!(!ink(&img, strip), "nothing left of the plot (deleted {}, tickLblPos {})", hidden.deleted, hidden.tick_labels);
        assert_columns_placed(&img, false);
    }
}

/// The light grey category axis line along zero (y 340), looked for between the first two
/// columns.
const BASELINE: (u32, u32, u32, u32) = (256, 336, 304, 344);
/// Under the first column, where its label Q1 goes.
const Q1_BELOW: (u32, u32, u32, u32) = (180, 393, 220, 410);

/// A deleted category axis (the add-in's charts carry their own label text boxes) draws neither
/// its labels nor its line. With only its labels off (`c:tickLblPos none`) the line stays.
#[test]
fn a_deleted_category_axis_draws_no_labels_and_no_line() {
    let img = image(&deck(&column_chart(&VALUES, LAYOUT, SHOWN)));
    assert!(text(&img, Q1_BELOW) && ink(&img, BASELINE), "shown: labels and line");

    let p = deck(&column_chart(&VALUES, LAYOUT, Axes { deleted: true, cat_deleted: true, ..SHOWN }));
    let c = chart_of(&p);
    assert!(c.cat_hidden && c.cat_deleted);
    let img = image(&p);
    assert!(!ink(&img, (PLOT.0 as u32, PLOT.3 as u32 + 2, PLOT.2 as u32, 479)), "nothing under the plot");
    assert!(!ink(&img, BASELINE), "no axis line");
    assert_columns_placed(&img, false);

    let p = deck(&column_chart(&VALUES, LAYOUT, Axes { cat_labels: "none", ..SHOWN }));
    let c = chart_of(&p);
    assert!(c.cat_hidden && !c.cat_deleted);
    let img = image(&p);
    assert!(!text(&img, Q1_BELOW) && ink(&img, BASELINE), "labels off, line kept");
    assert_columns_placed(&img, false);
}

/// Laid out automatically, a chart whose category axis is deleted keeps no empty band for its
/// labels: the plot reaches further down, as the value axis' gutter already did.
#[test]
fn a_deleted_category_axis_leaves_no_label_band() {
    let values = [12.0, 30.0, 6.0, 44.0];
    let fixed = Axes { scaling: r#"<c:max val="44"/><c:min val="0"/>"#, ..SHOWN };
    let bottom = |ax: Axes<'_>| {
        let got = bars(&image(&deck(&column_chart(&values, "<c:layout/>", ax))));
        assert_eq!(got.len(), 4, "{got:?}");
        assert!(got.iter().all(|b| b.3 == got[0].3), "the columns stand on one line: {got:?}");
        got[0].3
    };
    // The chart's area ends at 96 % of the frame (30 + 432 = 462); the 10 pt labels' band takes
    // 1.8 × 10 above it, a hidden axis only a 0.5 × 10 margin.
    assert!(near(bottom(fixed), 444.0), "{}", bottom(fixed));
    assert!(near(bottom(Axes { cat_deleted: true, ..fixed }), 457.0), "{}", bottom(Axes { cat_deleted: true, ..fixed }));
}

/// The XML of the one chart part in a saved deck.
fn chart_part(bytes: &[u8]) -> String {
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).expect("zip");
    let names: Vec<String> = zip.file_names().filter(|n| n.starts_with("ppt/charts/chart") && n.ends_with(".xml")).map(str::to_string).collect();
    assert_eq!(names.len(), 1, "{names:?}");
    let mut xml = String::new();
    zip.by_name(&names[0]).expect("chart part").read_to_string(&mut xml).expect("utf-8");
    xml
}

/// Editing the chart writes its part anew: the deleted axes stay deleted and the plot keeps its
/// manual layout, so overlays still line up after a save.
#[test]
fn an_edited_chart_keeps_its_deleted_axes_and_its_layout() {
    let mut p = deck(&column_chart(&VALUES, LAYOUT, Axes { deleted: true, cat_deleted: true, ..SHOWN }));
    let slide = Arc::make_mut(&mut p.slides[0]);
    let c = slide.shapes.iter_mut().find_map(|s| if let ShapeKind::Chart(c) = &mut s.kind { Some(c) } else { None }).expect("a chart");
    c.series[0].values[0] = Some(20.0);
    let bytes = deckcraft_pptx::export(&p).expect("export");
    let xml = chart_part(&bytes);
    assert!(!xml.contains(r#"<c:axId val="11"/>"#), "the part was written anew");
    assert!(xml.contains(r#"<c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/>"#), "{xml}");
    assert_eq!(xml.matches(r#"<c:delete val="1"/>"#).count(), 2, "{xml}");

    let back = deckcraft_pptx::import(&bytes).expect("import");
    let c = chart_of(&back);
    assert!(c.y_deleted && c.y_hidden && c.cat_deleted && c.cat_hidden);
    assert_eq!(c.plot_layout, Some(PlotLayout { x: 0.1, y: 0.2, w: 0.8, h: 0.6 }));
    let img = image(&back);
    let got = bars(&img);
    assert_eq!(got.len(), 4, "{got:?}");
    assert!(near(got[0].1, value_y(20.0, RANGE, false)), "the edited column: {:?}", got[0]);
    for (i, (&(l, t, r, _), &v)) in got.iter().zip(&VALUES).enumerate().skip(1) {
        let (want_l, want_r) = column_x(i, VALUES.len(), 60.0);
        let want_t = value_y(v.max(0.0), RANGE, false);
        assert!(near(l, want_l) && near(r, want_r) && near(t, want_t), "column {i}: {l}..{r}, top {t}");
    }
    let strip = (FRAME.0 as u32, PLOT.1 as u32 - 10, PLOT.0 as u32 - 1, PLOT.3 as u32 + 10);
    assert!(!ink(&img, strip) && !ink(&img, BASELINE) && !text(&img, Q1_BELOW), "still nothing drawn for either axis");
}

/// The bars of a bar chart: each run of rows with red in it, as (left, top, right, bottom), right
/// and bottom exclusive, from the bottom of the slide up (the first category's bar first).
fn hbars(img: &Image) -> Vec<(u32, u32, u32, u32)> {
    let red_y = |y: u32| (0..img.width).any(|x| rgb(img, x, y) == RED);
    let mut out = vec![];
    let mut y = img.height;
    while y > 0 {
        if !red_y(y - 1) {
            y -= 1;
            continue;
        }
        let y1 = y;
        while y > 0 && red_y(y - 1) {
            y -= 1;
        }
        let mid = (y + y1) / 2;
        let xs: Vec<u32> = (0..img.width).filter(|x| rgb(img, *x, mid) == RED).collect();
        out.push((xs[0], y, xs[xs.len() - 1] + 1, y1));
    }
    out
}

/// The kit's bar charts: `va.visible = False` writes a bare `<c:delete/>`, which the schema
/// reads as true.
fn bare_delete(xml: &str) -> String {
    let out = xml.replace(r#"<c:delete val="1"/><c:axPos val="b"/>"#, r#"<c:delete/><c:axPos val="b"/>"#);
    assert_ne!(out, xml, "the value axis was deleted");
    out
}

/// A bar chart on the manual plot: slots of plot height ÷ categories from the bottom up, each bar
/// the slot ÷ (1 + gap width), running from zero to its value across the fixed −10..44 axis.
#[test]
fn a_manual_inner_layout_places_bars_by_powerpoints_arithmetic() {
    let p = deck(&bare_delete(&bar_chart(&VALUES, LAYOUT, Axes { deleted: true, ..SHOWN })));
    let c = chart_of(&p);
    assert_eq!(c.plot_layout, Some(PlotLayout { x: 0.1, y: 0.2, w: 0.8, h: 0.6 }));
    assert!(c.y_deleted && c.y_hidden && c.kind.horizontal());
    let img = image(&p);
    let got = hbars(&img);
    assert_eq!(got.len(), VALUES.len(), "{got:?}");
    let (x0, y0, x1, y1) = PLOT;
    let slot = (y1 - y0) / VALUES.len() as f64;
    let bar = slot / 1.6;
    let x_of = |v: f64| x0 + (v - RANGE.0) / (RANGE.1 - RANGE.0) * (x1 - x0);
    for (i, (&(l, t, r, b), &v)) in got.iter().zip(&VALUES).enumerate() {
        let mid = y1 - slot * (i as f64 + 0.5);
        assert!(near(t, mid - bar / 2.0) && near(b, mid + bar / 2.0), "bar {i}: y {t}..{b}, want {}..{}", mid - bar / 2.0, mid + bar / 2.0);
        let (a, z) = (x_of(0.0), x_of(v));
        assert!(near(l, a.min(z)) && near(r, a.max(z)), "bar {i} ({v}): x {l}..{r}, want {}..{}", a.min(z), a.max(z));
    }
    // The last bar (44, the fixed maximum) reaches the plot's right edge; nothing is drawn under
    // the plot, where the deleted value axis' labels would go.
    assert!(near(got[3].2, x1), "{:?}", got[3]);
    assert!(!ink(&img, (x0 as u32, y1 as u32 + 2, x1 as u32, 470)), "no value axis under the plot");
}

/// Laid out automatically, a bar chart's deleted value axis draws no labels under the plot (the
/// strategy kit's deck showed a 0.0–10.0 axis there) and gives its band back to the bars.
#[test]
fn a_bar_charts_deleted_value_axis_draws_nothing_and_leaves_no_band() {
    let fixed = Axes { scaling: r#"<c:max val="44"/><c:min val="0"/>"#, ..SHOWN };
    let values = [12.0, 30.0, 6.0, 44.0];
    // The chart's area ends at 96 % of the frame (30 + 432 = 462). The 10 pt value labels take
    // 1.6 × 10 above that; a hidden axis keeps only a 0.5 × 10 margin. The area starts at 48, and
    // each of the four slots has a bar 1/1.6 of it, centred.
    let lowest_bottom = |plot_bottom: f64| {
        let slot = (plot_bottom - 48.0) / 4.0;
        plot_bottom - (slot - slot / 1.6) / 2.0
    };
    let img = image(&deck(&bar_chart(&values, "<c:layout/>", fixed)));
    let got = hbars(&img);
    assert!(near(got[0].3, lowest_bottom(446.0)), "{got:?}");
    assert!(text(&img, (got[0].0, 447, 830, 462)), "a shown axis has its labels under the plot");

    let p = deck(&bare_delete(&bar_chart(&values, "<c:layout/>", Axes { deleted: true, ..fixed })));
    assert!(chart_of(&p).y_deleted);
    let img = image(&p);
    let got = hbars(&img);
    assert_eq!(got.len(), 4, "{got:?}");
    assert!(near(got[0].3, lowest_bottom(457.0)), "{got:?}");
    assert!(!ink(&img, (FRAME.0 as u32 + 1, 458, (FRAME.0 + FRAME.2) as u32 - 1, 479)), "nothing under the plot");
}

/// `xml` with its categories Q1, Q2… renamed `names`, in order.
fn named(xml: &str, names: &[&str]) -> String {
    names.iter().enumerate().fold(xml.to_string(), |x, (i, n)| x.replace(&format!("<c:v>Q{}</c:v>", i + 1), &format!("<c:v>{n}</c:v>")))
}

/// The automatic layout's area: 4 % of the frame in from each side, (72, 48) to (808, 462).
const AREA: (f64, f64, f64, f64) = (72.0, 48.0, 808.0, 462.0);

/// The strategy kit's driver-tree bars, on a 0..12 axis that is deleted.
const BAR_VALUES: [f64; 4] = [2.0, 3.0, 8.0, 12.0];
const ZERO_TO_12: Axes<'static> = Axes { scaling: r#"<c:max val="12"/><c:min val="0"/>"#, deleted: true, ..SHOWN };

/// Laid out automatically, a bar chart's plot gives its category labels the room they need:
/// names such as "Derby and Kettering costs" are drawn in full inside the chart instead of being
/// cut off at its left edge, and short names leave the plot wider.
#[test]
fn bar_category_labels_get_the_room_they_need() {
    let names = ["Factory savings", "Derby and Kettering costs", "B2B sales", "Company-managed shops"];
    let img = image(&deck(&named(&bar_chart(&BAR_VALUES, "<c:layout/>", ZERO_TO_12), &names)));
    let got = hbars(&img);
    assert_eq!(got.len(), 4, "{got:?}");
    // The bars start at zero, the plot's left edge: right of the widest label (more than 100 pt
    // at 10 pt) and its 4 pt gap, and well short of the 40 % cap.
    let left = got[0].0;
    assert!(got.iter().all(|b| b.0 == left), "{got:?}");
    assert!(left as f64 > AREA.0 + 100.0 && (left as f64) < AREA.0 + 0.4 * (AREA.2 - AREA.0), "plot from {left}");
    // The widest label starts at the area's left edge; nothing is drawn left of it.
    assert!(text(&img, (AREA.0 as u32, 48, AREA.0 as u32 + 10, 462)), "the widest label starts at the area's edge");
    assert!(!ink(&img, (0, 0, AREA.0 as u32 - 2, 539)), "nothing runs off the chart");

    // Short names take only their own width: the plot reaches further left than the 4.5 em (45 pt)
    // band every bar chart had before.
    let short = hbars(&image(&deck(&bar_chart(&BAR_VALUES, "<c:layout/>", ZERO_TO_12))));
    assert!((short[0].0 as f64) < AREA.0 + 45.0 && short[0].0 < left, "{short:?}");
}

/// A label wider than 40 % of the chart wraps in a band that wide instead of squeezing the plot
/// further or running off the chart.
#[test]
fn bar_category_labels_past_the_cap_wrap() {
    let long = "Prices ahead, volume falls as the mix shifts to the cheapest lines in every region";
    let img = image(&deck(&named(&bar_chart(&BAR_VALUES, "<c:layout/>", ZERO_TO_12), &["Base case", long, "B2B", "Low case"])));
    let got = hbars(&img);
    assert_eq!(got.len(), 4, "{got:?}");
    let cap = AREA.0 + 0.4 * (AREA.2 - AREA.0);
    assert!(near(got[0].0, cap), "the plot starts at the cap ({cap}): {got:?}");
    assert!(!ink(&img, (0, 0, AREA.0 as u32 - 2, 539)), "nothing runs off the chart");
    // The long name (the second category from the bottom) is on more than one 10 pt line around
    // its bar's centre: one line's ink is about 10 px high, two lines' about 20.
    let mid = (got[1].1 + got[1].3) / 2;
    let rows: Vec<u32> = (mid - 40..mid + 40).filter(|y| text(&img, (AREA.0 as u32, *y, cap as u32 - 5, *y))).collect();
    let (top, bottom) = (rows[0], rows[rows.len() - 1]);
    assert!(bottom - top > 14, "wrapped onto two lines or more: rows {top}..{bottom}");
}

/// Beside a manual plot the labels have the room between the chart's edge and the plot; a name
/// wider than that wraps there rather than running off the chart.
#[test]
fn bar_category_labels_beside_a_manual_plot_wrap_within_the_chart() {
    let names = ["Factory savings", "Derby and Kettering costs", "B2B sales", "Company-managed shops"];
    let img = image(&deck(&named(&bar_chart(&BAR_VALUES, LAYOUT, ZERO_TO_12), &names)));
    let got = hbars(&img);
    assert!(got.iter().all(|b| near(b.0, PLOT.0)), "the bars start at the manual plot's left edge: {got:?}");
    assert!(text(&img, (FRAME.0 as u32, PLOT.1 as u32, PLOT.0 as u32 - 2, PLOT.3 as u32)), "labels left of the plot");
    assert!(!ink(&img, (0, 0, FRAME.0 as u32 - 2, 539)), "nothing left of the chart frame");
}

/// The strategy kit's null-point check: 57, (none), 200, 0, 288. Its cache has no point 1.
const WITH_NULL: [f64; 5] = [57.0, 999.0, 200.0, 0.0, 288.0];
const ZERO_TO_300: Axes<'static> = Axes { scaling: r#"<c:max val="300"/><c:min val="0"/>"#, deleted: true, ..SHOWN };

/// `xml` without its value for category `i`, as python-pptx writes a `None`.
fn without_point(xml: &str, i: usize, v: f64) -> String {
    let out = xml.replace(&format!(r#"<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>"#), "");
    assert_ne!(out, xml, "point {i} removed");
    out
}

/// The series' values labelled, `#,##0`, as the kit's charts are.
fn labelled(xml: &str) -> String {
    let labels = r##"<c:dLbls><c:numFmt formatCode="#,##0" sourceLinked="0"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>"##;
    xml.replacen(r#"<c:invertIfNegative val="0"/><c:cat>"#, &format!(r#"<c:invertIfNegative val="0"/>{labels}<c:cat>"#), 1)
}

/// Slot `i` of five across the manual plot: its centre.
fn slot5(i: usize) -> f64 {
    PLOT.0 + (PLOT.2 - PLOT.0) / 5.0 * (i as f64 + 0.5)
}

/// A point missing from the value cache is a gap: no column and no "0" label, while an explicit
/// zero keeps its label (DeckCraft previews showed the strategy kit's missing year as 0.0).
#[test]
fn a_missing_column_value_is_a_gap_and_zero_is_a_point() {
    let p = deck(&labelled(&without_point(&column_chart(&WITH_NULL, LAYOUT, ZERO_TO_300), 1, 999.0)));
    let c = chart_of(&p);
    assert_eq!(c.series[0].values, [Some(57.0), None, Some(200.0), Some(0.0), Some(288.0)]);
    let img = image(&p);
    let got = bars(&img);
    let centres: Vec<u32> = got.iter().map(|b| (b.0 + b.2) / 2).collect();
    assert_eq!(got.len(), 3, "57, 200 and 288 have columns: {got:?}");
    for (c, i) in centres.iter().zip([0, 2, 4]) {
        assert!(near(*c, slot5(i)), "column centres {centres:?}");
    }
    let above = |i: usize, top: u32| (slot5(i) as u32 - 25, top, slot5(i) as u32 + 25, PLOT.3 as u32 - 1);
    assert!(!text(&img, above(1, PLOT.1 as u32)), "no label in the missing year's slot");
    assert!(text(&img, above(3, 360)), "the explicit zero is labelled");
    assert!(text(&img, above(0, 300)), "57 is labelled");
}

/// A line chart of the same values, one 3 pt red line, `c:dispBlanksAs` set to `mode` (or left
/// out).
fn line_chart(mode: Option<&str>) -> String {
    let xml = without_point(&column_chart(&WITH_NULL, LAYOUT, ZERO_TO_300), 1, 999.0)
        .replace(r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>"#, r#"<c:lineChart><c:grouping val="standard"/>"#)
        .replace("</c:barChart>", "</c:lineChart>")
        .replace(r#"<c:invertIfNegative val="0"/>"#, "")
        .replace(r#"<c:gapWidth val="60"/>"#, "")
        .replace(
            r#"<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr>"#,
            r#"<c:spPr><a:ln w="38100"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>"#,
        );
    let blanks = mode.map(|m| format!(r#"<c:dispBlanksAs val="{m}"/>"#)).unwrap_or_default();
    xml.replace(r#"<c:plotVisOnly val="1"/>"#, &format!(r#"<c:plotVisOnly val="1"/>{blanks}"#))
}

/// Red line ink (antialiased) in the box (inclusive).
fn red_line(img: &Image, (x0, y0, x1, y1): (u32, u32, u32, u32)) -> bool {
    (y0..=y1).any(|y| (x0..=x1).any(|x| matches!(rgb(img, x, y), [r, g, b] if r > 200 && g < 90 && b < 90)))
}

/// A line breaks at a missing point (`gap`, also with no `c:dispBlanksAs`), joins the points
/// either side (`span`), or dips to zero (`zero`).
#[test]
fn a_line_breaks_spans_or_drops_to_zero_at_a_missing_point() {
    // 5 px to 9 units: 57 at y 338.7, 200 at 210, 0 at 390, 288 at 130.8.
    let y = |v: f64| PLOT.3 - v / 300.0 * (PLOT.3 - PLOT.1);
    let at = |x: f64, y: f64| (x as u32 - 2, y as u32 - 6, x as u32 + 2, y as u32 + 6);
    // Between the first point and the third, where only a line through the missing one goes.
    let gap = (slot5(0) as u32 + 8, PLOT.1 as u32, slot5(2) as u32 - 8, PLOT.3 as u32 + 2);
    for mode in [None, Some("gap")] {
        let p = deck(&line_chart(mode));
        assert_eq!(chart_of(&p).blanks, deckcraft_model::chart::Blanks::Gap);
        let img = image(&p);
        assert!(!red_line(&img, gap), "{mode:?}: the line breaks at the missing point");
        assert!(red_line(&img, at((slot5(2) + slot5(3)) / 2.0, (y(200.0) + y(0.0)) / 2.0)), "{mode:?}: the rest is drawn");
    }
    let img = image(&deck(&line_chart(Some("span"))));
    assert!(red_line(&img, at(slot5(1), (y(57.0) + y(200.0)) / 2.0)), "span: straight from 57 to 200");
    assert!(!red_line(&img, at(slot5(1), y(0.0) - 6.0)), "span: not down to zero");
    let img = image(&deck(&line_chart(Some("zero"))));
    assert!(red_line(&img, at(slot5(1), y(0.0) - 4.0)), "zero: down to the zero line");
    assert!(red_line(&img, at((slot5(0) + slot5(1)) / 2.0, (y(57.0) + y(0.0)) / 2.0)), "zero: from 57");
}
