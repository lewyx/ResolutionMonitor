#!/bin/bash
set -e

# Ensure we're in the repo root
cd "$(git rev-parse --show-toplevel)"

# Check for uncommitted changes
if ! git diff-index --quiet HEAD --; then
    echo "ERROR: There are uncommitted changes. Commit or stash them first."
    exit 1
fi

# Show current HEAD
echo "Current HEAD:"
git log -1 --oneline
echo

# Prompt for tag
read -p "Tag (e.g. v1.0.0): " tag
if [ -z "$tag" ]; then
    echo "ERROR: Tag cannot be empty."
    exit 1
fi

# Check tag doesn't already exist
if git rev-parse "$tag" >/dev/null 2>&1; then
    echo "ERROR: Tag '$tag' already exists."
    exit 1
fi

# Prompt for title (default to tag)
read -p "Title [$tag]: " title
title="${title:-$tag}"

# Prompt for release notes
read -p "Notes: " notes

# Confirm
echo
echo "Will create release:"
echo "  Tag:   $tag"
echo "  Title: $title"
echo "  Notes: $notes"
echo
read -p "Proceed? (y/N): " confirm
if [ "$confirm" != "y" ] && [ "$confirm" != "Y" ]; then
    echo "Aborted."
    exit 0
fi

# Create tag and release
git tag "$tag"
git push origin "$tag"
gh release create "$tag" --title "$title" --notes "$notes"

echo
echo "Release $tag created successfully."
