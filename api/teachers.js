export default function handler(req, res) {
  try {
    const list = JSON.parse(process.env.TEACHER_LIST || "[]");
    res.setHeader("Cache-Control", "s-maxage=60");
    res.status(200).json(list);
  } catch (e) {
    res.status(500).json({ error: "TEACHER_LIST env var is not valid JSON" });
  }
}
