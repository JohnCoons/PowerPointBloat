#!/usr/bin/env python3
"""
PowerPoint Cleanup Tool
Identifies and removes unused images, slide masters, and layouts from .pptx files.

Usage:
    python pptx_cleanup.py <path_to_unzipped_pptx> [--remove-images] [--remove-masters] [--remove-layouts]
"""

from __future__ import annotations

import argparse
import os
import shutil
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime
from pathlib import Path


class PPTXCleaner:
    """Analyzes and cleans unused content from PowerPoint presentations."""

    def __init__(self, pptx_folder: str | Path, verbose: bool = True) -> None:
        self.pptx_folder = Path(pptx_folder).resolve()
        self.verbose = verbose
        self.namespaces: dict[str, str] = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
            'ct': 'http://schemas.openxmlformats.org/package/2006/content-types'
        }

        # Register namespaces
        for prefix, uri in self.namespaces.items():
            ET.register_namespace(prefix, uri)

        # Active components
        self.active_slides: list[Path] = []
        self.active_masters: set[Path] = set()
        self.active_layouts: set[Path] = set()

        # All components
        self.all_masters: set[Path] = set()
        self.all_layouts: set[Path] = set()
        self.all_images: set[str] = set()

        # Referenced media
        self.image_references: dict[str, list[str]] = defaultdict(list)

        # Unused components
        self.unused_masters: set[Path] = set()
        self.unused_layouts: set[Path] = set()
        self.unused_images: set[str] = set()

    def log(self, message: str) -> None:
        """Print message if verbose mode enabled."""
        if self.verbose:
            print(message)

    def validate_folder(self) -> bool:
        """Validate this is an unzipped PowerPoint folder."""
        required = [
            self.pptx_folder / 'ppt' / 'presentation.xml',
            self.pptx_folder / '[Content_Types].xml'
        ]
        for path in required:
            if not path.exists():
                print(f"Error: {path} not found. Is this an unzipped .pptx?")
                return False
        return True

    def parse_presentation_structure(self) -> bool:
        """Parse presentation.xml to find active slides and masters."""
        self.log("\n[1/5] Parsing presentation structure...")

        pres_file = self.pptx_folder / 'ppt' / 'presentation.xml'
        pres_rels = self.pptx_folder / 'ppt' / '_rels' / 'presentation.xml.rels'

        tree = ET.parse(pres_file)
        root = tree.getroot()

        # Get all master relationship IDs
        master_list = root.find('.//p:sldMasterIdLst', self.namespaces)
        all_master_rids = []
        if master_list is not None:
            for master_id in master_list.findall('.//p:sldMasterId', self.namespaces):
                rid = master_id.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if rid:
                    all_master_rids.append(rid)

        # Get active slide relationship IDs
        slide_list = root.find('.//p:sldIdLst', self.namespaces)
        active_slide_rids = []
        if slide_list is not None:
            for slide_id in slide_list.findall('.//p:sldId', self.namespaces):
                rid = slide_id.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if rid:
                    active_slide_rids.append(rid)

        # Resolve relationships
        rels_tree = ET.parse(pres_rels)
        rels_root = rels_tree.getroot()

        # Map all masters
        for rid in all_master_rids:
            rel = rels_root.find(f".//rel:Relationship[@Id='{rid}']", self.namespaces)
            if rel is not None:
                target = rel.get('Target')
                if target:
                    master_path = (self.pptx_folder / 'ppt' / target).resolve()
                    self.all_masters.add(master_path)

        # Find active slides and their masters
        for rid in active_slide_rids:
            rel = rels_root.find(f".//rel:Relationship[@Id='{rid}']", self.namespaces)
            if rel is not None:
                target = rel.get('Target')
                if target:
                    slide_path = (self.pptx_folder / 'ppt' / target).resolve()
                    self.active_slides.append(slide_path)
                    master = self._find_master_for_slide(slide_path)
                    if master:
                        self.active_masters.add(master)

        self.log(f"    Found {len(self.active_slides)} active slides")
        self.log(f"    Found {len(self.active_masters)}/{len(self.all_masters)} masters in use")
        return True

    def _find_master_for_slide(self, slide_path: Path) -> Path | None:
        """Find which master a slide uses via its layout."""
        slide_rels = slide_path.parent / '_rels' / f"{slide_path.name}.rels"
        if not slide_rels.exists():
            return None

        try:
            tree = ET.parse(slide_rels)
            root = tree.getroot()

            for rel in root.findall('.//rel:Relationship', self.namespaces):
                if 'slideLayout' in rel.get('Type', ''):
                    target = rel.get('Target')
                    if target:
                        layout_path = (slide_path.parent / target).resolve()
                        self.active_layouts.add(layout_path)
                        return self._find_master_for_layout(layout_path)
        except ET.ParseError:
            pass
        return None

    def _find_master_for_layout(self, layout_path: Path) -> Path | None:
        """Find which master a layout belongs to."""
        layout_rels = layout_path.parent / '_rels' / f"{layout_path.name}.rels"
        if not layout_rels.exists():
            return None

        try:
            tree = ET.parse(layout_rels)
            root = tree.getroot()

            for rel in root.findall('.//rel:Relationship', self.namespaces):
                if 'slideMaster' in rel.get('Type', ''):
                    target = rel.get('Target')
                    if target:
                        return (layout_path.parent / target).resolve()
        except ET.ParseError:
            pass
        return None

    def find_all_layouts(self) -> None:
        """Find all layouts and identify which are used by active masters."""
        self.log("\n[2/5] Analyzing layouts...")

        layouts_dir = self.pptx_folder / 'ppt' / 'slideLayouts'
        if layouts_dir.exists():
            for layout_file in layouts_dir.glob('slideLayout*.xml'):
                self.all_layouts.add(layout_file.resolve())

        # Find layouts used by active masters
        for master_path in self.active_masters:
            master_rels = master_path.parent / '_rels' / f"{master_path.name}.rels"
            if not master_rels.exists():
                continue

            try:
                tree = ET.parse(master_rels)
                root = tree.getroot()

                for rel in root.findall('.//rel:Relationship', self.namespaces):
                    if 'slideLayout' in rel.get('Type', ''):
                        target = rel.get('Target')
                        if target:
                            layout_path = (master_path.parent / target).resolve()
                            self.active_layouts.add(layout_path)
            except ET.ParseError:
                pass

        self.log(f"    Found {len(self.active_layouts)}/{len(self.all_layouts)} layouts in use")

    def scan_media_files(self) -> None:
        """Scan for all media files."""
        self.log("\n[3/5] Scanning media files...")

        media_path = self.pptx_folder / 'ppt' / 'media'
        if not media_path.exists():
            self.log("    No media folder found")
            return

        image_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.emf', '.wmf', '.svg'}
        for file in media_path.iterdir():
            if file.is_file() and file.suffix.lower() in image_extensions:
                self.all_images.add(file.name)

        self.log(f"    Found {len(self.all_images)} image files")

    def find_referenced_media(self) -> None:
        """Find all media referenced by active components."""
        self.log("\n[4/5] Finding media references...")

        components = list(self.active_slides) + list(self.active_masters) + list(self.active_layouts)

        for component in components:
            rels_file = component.parent / '_rels' / f"{component.name}.rels"
            if not rels_file.exists():
                continue

            try:
                tree = ET.parse(rels_file)
                root = tree.getroot()

                for rel in root.findall('.//rel:Relationship', self.namespaces):
                    target = rel.get('Target', '')
                    if 'media/' in target or '../media/' in target:
                        filename = os.path.basename(target)
                        self.image_references[filename].append(str(component.name))
            except ET.ParseError:
                pass

        self.log(f"    Found {len(self.image_references)} images referenced by active content")

    def calculate_unused(self) -> None:
        """Calculate all unused components."""
        self.log("\n[5/5] Calculating unused components...")

        self.unused_masters = self.all_masters - self.active_masters
        self.unused_layouts = self.all_layouts - self.active_layouts
        self.unused_images = self.all_images - set(self.image_references.keys())

        self.log(f"    Unused masters: {len(self.unused_masters)}")
        self.log(f"    Unused layouts: {len(self.unused_layouts)}")
        self.log(f"    Unused images: {len(self.unused_images)}")

    def generate_report(self) -> int:
        """Generate cleanup report. Returns total bytes of unused images."""
        print("\n" + "="*70)
        print("CLEANUP REPORT")
        print("="*70)

        print(f"\nActive slides: {len(self.active_slides)}")
        print(f"Active masters: {len(self.active_masters)} / {len(self.all_masters)}")
        print(f"Active layouts: {len(self.active_layouts)} / {len(self.all_layouts)}")
        print(f"Referenced images: {len(self.image_references)} / {len(self.all_images)}")

        # Calculate space savings
        total_size = 0
        if self.unused_images:
            print(f"\n--- UNUSED IMAGES ({len(self.unused_images)}) ---")
            for img in sorted(self.unused_images):
                img_path = self.pptx_folder / 'ppt' / 'media' / img
                size = img_path.stat().st_size if img_path.exists() else 0
                total_size += size
                print(f"  {img}: {size:,} bytes")
            print(f"\nTotal space to reclaim: {total_size:,} bytes ({total_size/1024/1024:.2f} MB)")

        if self.unused_masters:
            print(f"\n--- UNUSED MASTERS ({len(self.unused_masters)}) ---")
            for master in sorted(self.unused_masters):
                print(f"  {master.name}")

        if self.unused_layouts:
            print(f"\n--- UNUSED LAYOUTS ({len(self.unused_layouts)}) ---")
            count = min(20, len(self.unused_layouts))
            for layout in sorted(list(self.unused_layouts))[:count]:
                print(f"  {layout.name}")
            if len(self.unused_layouts) > 20:
                print(f"  ... and {len(self.unused_layouts) - 20} more")

        return total_size

    def create_backup(self) -> Path:
        """Create backup of critical files."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_dir = self.pptx_folder / f"backup_{timestamp}"
        backup_dir.mkdir(exist_ok=True)

        files = ['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml']
        for file_path in files:
            src = self.pptx_folder / file_path
            if src.exists():
                dst = backup_dir / file_path
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst)

        print(f"\nBackup created: {backup_dir.name}")
        return backup_dir

    def remove_unused_images(self) -> int:
        """Remove unused images from ppt/media/. Returns count of removed files."""
        if not self.unused_images:
            print("No unused images to remove.")
            return 0

        removed = 0
        for img in self.unused_images:
            img_path = self.pptx_folder / 'ppt' / 'media' / img
            if img_path.exists():
                img_path.unlink()
                removed += 1
                print(f"  Removed: {img}")

        print(f"\nRemoved {removed} unused images.")
        return removed

    def remove_unused_masters(self) -> int:
        """Remove unused masters and update XML files. Returns count of removed masters."""
        if not self.unused_masters:
            print("No unused masters to remove.")
            return 0

        self.create_backup()

        # Find relationship IDs for unused masters
        pres_rels = self.pptx_folder / 'ppt' / '_rels' / 'presentation.xml.rels'
        rels_tree = ET.parse(pres_rels)
        rels_root = rels_tree.getroot()

        rids_to_delete = []
        for master in self.unused_masters:
            master_target = f"slideMasters/{master.name}"
            for rel in rels_root.findall('.//rel:Relationship', self.namespaces):
                if master_target in rel.get('Target', ''):
                    rids_to_delete.append(rel.get('Id'))
                    break

        # Update presentation.xml
        pres_file = self.pptx_folder / 'ppt' / 'presentation.xml'
        pres_tree = ET.parse(pres_file)
        pres_root = pres_tree.getroot()

        master_list = pres_root.find('.//p:sldMasterIdLst', self.namespaces)
        if master_list is not None:
            for master_id in list(master_list.findall('.//p:sldMasterId', self.namespaces)):
                rid = master_id.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if rid in rids_to_delete:
                    master_list.remove(master_id)

        pres_tree.write(pres_file, encoding='UTF-8', xml_declaration=True)

        # Update [Content_Types].xml
        ct_file = self.pptx_folder / '[Content_Types].xml'
        ct_tree = ET.parse(ct_file)
        ct_root = ct_tree.getroot()

        for master in self.unused_masters:
            part_name = f"/ppt/slideMasters/{master.name}"
            for override in list(ct_root.findall('.//ct:Override', self.namespaces)):
                if override.get('PartName') == part_name:
                    ct_root.remove(override)

        ct_tree.write(ct_file, encoding='UTF-8', xml_declaration=True)

        # Delete master files
        removed = 0
        for master in self.unused_masters:
            if master.exists():
                master.unlink()
                removed += 1
            rels = master.parent / '_rels' / f"{master.name}.rels"
            if rels.exists():
                rels.unlink()

        print(f"\nRemoved {removed} unused masters.")
        return removed

    def remove_unused_layouts(self) -> int:
        """Remove unused layouts and update XML files. Returns count of removed layouts."""
        if not self.unused_layouts:
            print("No unused layouts to remove.")
            return 0

        self.create_backup()

        # Update [Content_Types].xml
        ct_file = self.pptx_folder / '[Content_Types].xml'
        ct_tree = ET.parse(ct_file)
        ct_root = ct_tree.getroot()

        for layout in self.unused_layouts:
            part_name = f"/ppt/slideLayouts/{layout.name}"
            for override in list(ct_root.findall('.//ct:Override', self.namespaces)):
                if override.get('PartName') == part_name:
                    ct_root.remove(override)

        ct_tree.write(ct_file, encoding='UTF-8', xml_declaration=True)

        # Delete layout files
        removed = 0
        for layout in self.unused_layouts:
            if layout.exists():
                layout.unlink()
                removed += 1
                print(f"  Removed: {layout.name}")
            rels = layout.parent / '_rels' / f"{layout.name}.rels"
            if rels.exists():
                rels.unlink()

        print(f"\nRemoved {removed} unused layouts.")
        return removed

    def save_removal_scripts(self) -> None:
        """Generate bash scripts for removal."""
        # Image removal script
        script = self.pptx_folder / 'remove_unused_images.sh'
        with open(script, 'w', newline='\n') as f:
            f.write("#!/bin/bash\n# Remove unused images\n\n")
            if self.unused_images:
                f.write(f"echo 'Removing {len(self.unused_images)} unused images...'\n")
                for img in sorted(self.unused_images):
                    f.write(f"rm -f 'ppt/media/{img}'\n")
                f.write("\necho 'Done!'\n")
            else:
                f.write("echo 'No unused images.'\n")

        # Master removal script
        script = self.pptx_folder / 'remove_unused_masters.sh'
        with open(script, 'w', newline='\n') as f:
            f.write("#!/bin/bash\n# Remove unused masters\n")
            f.write("# WARNING: Also requires updating presentation.xml and [Content_Types].xml\n\n")
            if self.unused_masters:
                f.write(f"echo 'Removing {len(self.unused_masters)} unused masters...'\n")
                for master in sorted(self.unused_masters):
                    rel = master.relative_to(self.pptx_folder)
                    f.write(f"rm -f '{rel}'\n")
                    f.write(f"rm -f '{rel.parent}/_rels/{master.name}.rels'\n")
                f.write("\necho 'Done! Update [Content_Types].xml and presentation.xml manually.'\n")
            else:
                f.write("echo 'No unused masters.'\n")

        # Layout removal script
        script = self.pptx_folder / 'remove_unused_layouts.sh'
        with open(script, 'w', newline='\n') as f:
            f.write("#!/bin/bash\n# Remove unused layouts\n")
            f.write("# WARNING: Also requires updating [Content_Types].xml\n\n")
            if self.unused_layouts:
                f.write(f"echo 'Removing {len(self.unused_layouts)} unused layouts...'\n")
                for layout in sorted(self.unused_layouts):
                    rel = layout.relative_to(self.pptx_folder)
                    f.write(f"rm -f '{rel}'\n")
                    f.write(f"rm -f '{rel.parent}/_rels/{layout.name}.rels'\n")
                f.write("\necho 'Done! Update [Content_Types].xml manually.'\n")
            else:
                f.write("echo 'No unused layouts.'\n")

        # Text list
        with open(self.pptx_folder / 'unused_components.txt', 'w') as f:
            f.write("UNUSED COMPONENTS\n" + "="*70 + "\n\n")

            f.write(f"UNUSED IMAGES ({len(self.unused_images)}):\n")
            for img in sorted(self.unused_images):
                f.write(f"  ppt/media/{img}\n")

            f.write(f"\nUNUSED MASTERS ({len(self.unused_masters)}):\n")
            for m in sorted(self.unused_masters):
                f.write(f"  {m.relative_to(self.pptx_folder)}\n")

            f.write(f"\nUNUSED LAYOUTS ({len(self.unused_layouts)}):\n")
            for l in sorted(self.unused_layouts):
                f.write(f"  {l.relative_to(self.pptx_folder)}\n")

        print("\nGenerated scripts:")
        print("  - remove_unused_images.sh")
        print("  - remove_unused_masters.sh")
        print("  - remove_unused_layouts.sh")
        print("  - unused_components.txt")

    def analyze(self) -> bool:
        """Run full analysis."""
        if not self.validate_folder():
            return False

        self.parse_presentation_structure()
        self.find_all_layouts()
        self.scan_media_files()
        self.find_referenced_media()
        self.calculate_unused()
        self.generate_report()
        self.save_removal_scripts()
        return True


def main() -> None:
    parser = argparse.ArgumentParser(
        description='PowerPoint Cleanup Tool - Remove unused images, masters, and layouts',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python pptx_cleanup.py ./mypresentation/          # Analyze only
  python pptx_cleanup.py ./mypresentation/ --remove-images
  python pptx_cleanup.py ./mypresentation/ --remove-layouts
  python pptx_cleanup.py ./mypresentation/ --remove-masters
        """
    )
    parser.add_argument('folder', help='Path to unzipped .pptx folder')
    parser.add_argument('--remove-images', action='store_true', help='Remove unused images')
    parser.add_argument('--remove-layouts', action='store_true', help='Remove unused layouts')
    parser.add_argument('--remove-masters', action='store_true', help='Remove unused masters (advanced)')
    parser.add_argument('-q', '--quiet', action='store_true', help='Quiet mode')

    args = parser.parse_args()

    if not os.path.exists(args.folder):
        print(f"Error: Folder '{args.folder}' does not exist!")
        sys.exit(1)

    cleaner = PPTXCleaner(args.folder, verbose=not args.quiet)

    if not cleaner.analyze():
        sys.exit(1)

    if args.remove_images:
        print("\n" + "="*70)
        print("REMOVING UNUSED IMAGES")
        print("="*70)
        cleaner.remove_unused_images()

    if args.remove_layouts:
        print("\n" + "="*70)
        print("REMOVING UNUSED LAYOUTS")
        print("="*70)
        cleaner.remove_unused_layouts()

    if args.remove_masters:
        print("\n" + "="*70)
        print("REMOVING UNUSED MASTERS")
        print("="*70)
        cleaner.remove_unused_masters()

    if not args.remove_images and not args.remove_layouts and not args.remove_masters:
        print("\n" + "="*70)
        print("NEXT STEPS")
        print("="*70)
        print("\nTo remove unused content:")
        print("  1. Run with --remove-images to delete unused images")
        print("  2. Run with --remove-layouts to delete unused layouts")
        print("  3. Run with --remove-masters to delete unused masters (advanced)")
        print("  4. Re-zip folder and rename to .pptx")


if __name__ == '__main__':
    main()
