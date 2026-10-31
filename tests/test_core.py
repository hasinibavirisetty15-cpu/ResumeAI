"""Unit tests for the pure logic in resume_analyzer.py.

Run from the project folder with:  python -m unittest discover tests

The Streamlit UI code is skipped: only constants and functions are
loaded, so these tests need no Streamlit session.
"""
import ast
import os
import re
import unittest

SOURCE = os.path.join(
    os.path.dirname(__file__), "..", "resume_analyzer.py"
)


def load_core():
    with open(SOURCE, encoding="utf-8") as handle:
        tree = ast.parse(handle.read())

    keep = []

    for node in tree.body:
        if isinstance(node, ast.FunctionDef):
            keep.append(node)
        elif isinstance(node, ast.Assign) and all(
            isinstance(t, ast.Name) and t.id.isupper()
            for t in node.targets
        ):
            keep.append(node)

    module = ast.Module(body=keep, type_ignores=[])
    namespace = {"re": re}
    exec(compile(module, SOURCE, "exec"), namespace)
    return namespace


core = load_core()

JD = """# Junior Python Developer

**Company:** Example Tech (Fictional)

### About the Role
We are looking for a motivated Junior Python Developer to help build, test, and maintain reliable software.

### Required Skills
- Python
- SQL
- Git
- Object-oriented programming

### Preferred Skills
- Flask
- AWS
"""

RESUME = """ASHA RAO
asha@example.com | 9876543210

SKILLS
Python, C++, SQL, Git

PROJECTS
Library App | Python, SQL
- worked on a book search feature

EDUCATION
BSc Computer Science, Example University | 2027

CERTIFICATIONS
Python Course
"""


class NormalizeTests(unittest.TestCase):
    def test_node_js_is_not_doubled(self):
        self.assertEqual(core["normalize_text"]("Node.js and node"),
                         "node.js and node.js")

    def test_algorithms_not_mangled(self):
        self.assertIn("algorithms", core["normalize_text"]("Algorithms"))
        self.assertNotIn("algorithmss", core["normalize_text"]("Algorithms"))

    def test_c_not_found_inside_cpp(self):
        self.assertFalse(core["phrase_present"]("I know C++", "c"))
        self.assertTrue(core["phrase_present"]("I know C, Java", "c"))


class SkillTests(unittest.TestCase):
    def test_required_and_preferred_split(self):
        skills = core["extract_skills"](JD)
        required, preferred = core["detect_required_preferred_skills"](
            JD, skills
        )
        self.assertIn("python", required)
        self.assertIn("sql", required)
        self.assertIn("flask", preferred)
        self.assertIn("aws", preferred)
        self.assertNotIn("flask", required)

    def test_skill_match_score(self):
        result = core["skill_matching"](RESUME, JD)
        self.assertIn("python", result["matching_required"])
        self.assertIn("oop", result["missing_required"])
        self.assertTrue(0 <= result["score"] <= 100)

    def test_display_names(self):
        self.assertEqual(core["display_skill"]("oop"), "OOP")
        self.assertEqual(core["display_skill"]("sql"), "SQL")
        self.assertEqual(core["display_skill"]("python"), "Python")


class ExtractionTests(unittest.TestCase):
    def test_contact_details(self):
        self.assertEqual(core["extract_name"](RESUME), "ASHA RAO")
        self.assertEqual(core["extract_email"](RESUME), "asha@example.com")
        self.assertEqual(core["extract_phone"](RESUME), "9876543210")

    def test_target_role_is_a_clean_title(self):
        self.assertEqual(core["detect_target_role"](JD),
                         "Junior Python Developer")

    def test_target_role_from_sentence(self):
        text = "We are looking for a motivated Data Analyst to join our team."
        self.assertEqual(core["detect_target_role"](text), "Data Analyst")


class ScoreTests(unittest.TestCase):
    def test_overall_score_weights(self):
        self.assertEqual(core["calculate_overall_score"](100, 100, 100, 100), 100)
        self.assertEqual(core["calculate_overall_score"](0, 0, 0, 0), 0)

    def test_structure_detects_sections(self):
        result = core["analyze_structure"](RESUME)
        self.assertTrue(result["sections"]["Education"])
        self.assertTrue(result["sections"]["Projects"])
        self.assertFalse(result["sections"]["Achievements"])


class OptimizedResumeTests(unittest.TestCase):
    def build(self):
        data = {
            "resume_text": RESUME,
            "skill_data": core["skill_matching"](RESUME, JD),
            "target_role": core["detect_target_role"](JD),
        }
        return core["generate_optimized_resume"](data)

    def test_uses_real_resume_content(self):
        draft = self.build()
        self.assertIn("ASHA RAO", draft)
        self.assertIn("Library App", draft)
        self.assertIn("BSc Computer Science", draft)

    def test_no_placeholder_text(self):
        draft = self.build()
        self.assertNotIn("Add degree", draft)
        self.assertNotIn("Highlight projects", draft)
        self.assertNotIn("ALEX MORGAN", draft)

    def test_missing_section_is_reported(self):
        self.assertIn("Achievements", self.build())


if __name__ == "__main__":
    unittest.main()
