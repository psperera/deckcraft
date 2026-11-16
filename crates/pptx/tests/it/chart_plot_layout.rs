//! U30: native charts drawn with PowerPoint's plot geometry. A manual inner plot layout puts the
//! plot exactly where the file says, fixed axis ends are kept exactly, and a deleted value axis
//! draws nothing. Each chart is XML written here, imported and rendered; no Office file is used.

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

/// The value axis' settings that vary between the tests.
#[derive(Clone, Copy)]
struct ValAx<'a> {
    /// `c:scaling` children other than the orientation.
    scaling: &'a str,
    reversed: bool,
    deleted: bool,
    /// `c:tickLblPos`.
    tick_labels: &'a str,
    major: Option<f64>,
}

/// A clustered column chart of one red series, `values` in categories Q1, Q2…, with a gap
/// width of 60 %, 10 pt text and no title or legend.
fn column_chart(values: &[f64], layout: &str, ax: ValAx<'_>) -> String {
    let n = values.len();
    let pts: String = values.iter().enumerate().map(|(i, v)| format!(r#"<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>"#)).collect();
    let cats: String = (0..n).map(|i| format!(r#"<c:pt idx="{i}"><c:v>Q{}</c:v></c:pt>"#, i + 1)).collect();
    let orientation = if ax.reversed { "maxMin" } else { "minMax" };
    let ValAx { scaling, tick_labels, .. } = ax;
    let delete = u8::from(ax.deleted);
    let major = ax.major.map(|m| format!(r#"<c:majorUnit val="{m}"/>"#)).unwrap_or_default();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace {NS}><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>{layout}<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/><c:cat><c:strRef><c:strCache><c:ptCount val="{n}"/>{cats}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{n}"/>{pts}</c:numCache></c:numRef></c:val></c:ser><c:gapWidth val="60"/><c:axId val="11"/><c:axId val="12"/></c:barChart><c:catAx><c:axId val="11"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="12"/><c:crosses val="autoZero"/></c:catAx><c:valAx><c:axId val="12"/><c:scaling><c:orientation val="{orientation}"/>{scaling}</c:scaling><c:delete val="{delete}"/><c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="0"/><c:tickLblPos val="{tick_labels}"/><c:crossAx val="11"/><c:crosses val="autoZero"/><c:crossBetween val="between"/>{major}</c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:chartSpace>"#
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
    let ax = ValAx { scaling: FIXED, reversed: false, deleted: true, tick_labels: "nextTo", major: None };
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
    let ax = ValAx { scaling: FIXED, reversed: true, deleted: true, tick_labels: "nextTo", major: None };
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
    let ax = ValAx { scaling: r#"<c:max val="288"/><c:min val="0"/>"#, reversed: false, deleted: false, tick_labels: "nextTo", major: Some(50.0) };
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
    let shown = ValAx { scaling: FIXED, reversed: false, deleted: false, tick_labels: "nextTo", major: None };
    // Left of the plot, from above its top to below its bottom (the category labels are under
    // the columns, further right).
    let strip = (FRAME.0 as u32, PLOT.1 as u32 - 10, PLOT.0 as u32 - 1, PLOT.3 as u32 + 10);
    let img = image(&deck(&column_chart(&VALUES, LAYOUT, shown)));
    assert!(text(&img, strip), "a shown axis has its labels there");
    assert_columns_placed(&img, false);
    for hidden in [ValAx { deleted: true, ..shown }, ValAx { tick_labels: "none", ..shown }] {
        let p = deck(&column_chart(&VALUES, LAYOUT, hidden));
        assert!(chart_of(&p).y_hidden);
        let img = image(&p);
        assert!(!ink(&img, strip), "nothing left of the plot (deleted {}, tickLblPos {})", hidden.deleted, hidden.tick_labels);
        assert_columns_placed(&img, false);
    }
}
