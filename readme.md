Behavior (per directory):

1. Require files:

   - background.png
   - background_bar_header.png
   - background_bar_footer.png
   - productNameAttr.txt (5 lines: font, bg_color, text_color, font_size, name_top_left x,y)
   - productInfoAttr.txt (4 lines: font, text_color, font_size, info_top_left x,y)
   - positions.txt (3 lines: centers for B, A, C as x,y)
   - maxheight.txt (3 lines: A_max_h, B_max_h, C_max_h; positive integers)
     If any missing → warn and skip directory.

2. Validate attr files, font existence, hex colors, positive font sizes.

3. Use Products.xlsx (row 2+):
   A=Folder, B=SKU, C=ProductName, D=ProductInfo, E=A rot, F=B rot, G=C rot.
   - Verify {SKU}\_A.png / \_B.png / \_C.png exist in folder
   - Copy background.png → output/<Folder>/{SKU}\_{ProductNameNoSpaces}.png
   - Compose: draw ProductName (with BG), draw ProductInfo; paste A/B/C centered at
     positions from positions.txt with the specified rotations (deg; negative=left, positive=right).

Scaling rule:

- Read A/B/C max heights from maxheight.txt (lines 1/2/3 respectively). Images are scaled
  down to those max heights (keeping aspect) BEFORE rotation.

Output:

- All generated files go under: <root>/output/<same directory structure as source>.
  """
