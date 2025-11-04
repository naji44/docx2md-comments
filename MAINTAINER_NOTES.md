# Maintainer Notes

Internal notes for maintaining this repository.

## Making a Release

1. Update version in code/README if applicable
2. Go to [Releases](https://github.com/naji44/docx2md-comments/releases)
3. Click **Draft a new release**
4. Create tag: `v1.0.0` (or next version)
5. Release title: `v1.0.0 - [Brief description]`
6. Add release notes describing changes
7. Click **Publish release**

## Pre-release Checklist

- [ ] All tests pass (if you add tests)
- [ ] README is up to date
- [ ] Requirements.txt is current
- [ ] Registry files have correct paths
- [ ] License year is current

## Repository Structure
```
docx2md-comments/
├── src/
│   └── word_to_md_with_comments.py
├── registry/
│   ├── add_context_menu.reg
│   └── remove_context_menu.reg
├── examples/ (optional)
│   └── sample.docx
├── requirements.txt
├── LICENSE
├── README.md
└── .gitignore
```

## Testing Locally
```bash
# Test basic conversion
python src/word_to_md_with_comments.py examples/sample.docx

# Verify output
cat examples/sample_feedback.md
```

## Common Issues

**Registry file doesn't work:**
- Verify Python path in `.reg` file
- Run as Administrator
- Check backslashes are escaped (`\\`)

**Comments not appearing:**
- Ensure Word document has actual comments saved
- Check `word/comments.xml` exists in the .docx
```

---

## Other Improvements to Consider:

### 1. **Create an `examples/` folder**
```
examples/
├── README.md
└── sample.docx (a small test file with 1-2 comments)
