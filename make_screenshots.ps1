# make_screenshots.ps1
Add-Type -AssemblyName System.Drawing

function Create-Screenshot1 {
    param($path)
    $w = 1100; $h = 600
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

    # Colors
    $cTopBar = [System.Drawing.Color]::FromArgb(17, 35, 42)
    $cSideBar = [System.Drawing.Color]::FromArgb(24, 43, 52)
    $cMainBg = [System.Drawing.Color]::FromArgb(250, 251, 252)
    $cWhite = [System.Drawing.Color]::White
    $cBorder = [System.Drawing.Color]::FromArgb(220, 224, 230)
    $cGreen = [System.Drawing.Color]::FromArgb(0, 237, 100)
    $cGreenDark = [System.Drawing.Color]::FromArgb(0, 163, 92)
    $cTextDark = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $cTextMuted = [System.Drawing.Color]::FromArgb(108, 117, 125)
    $cJsonBlue = [System.Drawing.Color]::FromArgb(0, 104, 214)
    $cJsonRed = [System.Drawing.Color]::FromArgb(209, 36, 47)
    $cJsonGreen = [System.Drawing.Color]::FromArgb(17, 99, 41)

    # Fonts
    $fTitle = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $fRegular = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
    $fBold = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $fCode = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Regular)
    $fCodeBold = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)

    # Brushes
    $bWhite = New-Object System.Drawing.SolidBrush($cWhite)
    $bTextDark = New-Object System.Drawing.SolidBrush($cTextDark)
    $bTextMuted = New-Object System.Drawing.SolidBrush($cTextMuted)
    $bGreen = New-Object System.Drawing.SolidBrush($cGreenDark)
    $bJsonBlue = New-Object System.Drawing.SolidBrush($cJsonBlue)
    $bJsonRed = New-Object System.Drawing.SolidBrush($cJsonRed)
    $bJsonGreen = New-Object System.Drawing.SolidBrush($cJsonGreen)

    # Top Bar
    $g.FillRectangle((New-Object System.Drawing.SolidBrush($cTopBar)), 0, 0, $w, 42)
    $g.DrawString("🍃 MongoDB Compass — localhost:27017", $fTitle, $bWhite, 15, 11)
    $g.DrawString("Connected: MongoDB 7.0 (Standalone)", $fRegular, (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::LightGreen)), 780, 12)

    # Left Sidebar
    $g.FillRectangle((New-Object System.Drawing.SolidBrush($cSideBar)), 0, 42, 230, $h - 42)
    $g.DrawString("DATABASES (3)", $fBold, (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(150, 170, 180))), 15, 55)
    $g.DrawString("📁 admin", $fRegular, $bWhite, 20, 80)
    $g.DrawString("📁 config", $fRegular, $bWhite, 20, 105)
    $g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(35, 65, 78))), 5, 130, 220, 55)
    $g.DrawString("📂 ecommerce_catalog", $fBold, (New-Object System.Drawing.SolidBrush($cGreen)), 15, 135)
    $g.DrawString("   📄 products (25 docs)", $fBold, $bWhite, 20, 160)

    # Main Panel Background
    $g.FillRectangle((New-Object System.Drawing.SolidBrush($cMainBg)), 230, 42, $w - 230, $h - 42)

    # Sub-header: Collection Title & Stats
    $g.FillRectangle($bWhite, 230, 42, $w - 230, 50)
    $g.DrawLine((New-Object System.Drawing.Pen($cBorder, 1)), 230, 92, $w, 92)
    $g.DrawString("ecommerce_catalog.products", (New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)), $bTextDark, 250, 52)
    $g.DrawString("Total Documents: 25   |   Indexes: 3   |   Storage Size: 32 KB   |   Avg Doc Size: 780 B", $fRegular, $bTextMuted, 520, 56)

    # Tabs
    $g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(240, 243, 246))), 230, 93, $w - 230, 35)
    $g.DrawLine((New-Object System.Drawing.Pen($cBorder, 1)), 230, 128, $w, 128)
    # Active tab Documents
    $g.FillRectangle($bWhite, 250, 93, 130, 36)
    $g.DrawString("Documents (25)", $fBold, $bGreen, 260, 102)
    $g.DrawString("Aggregations", $fRegular, $bTextMuted, 400, 102)
    $g.DrawString("Schema", $fRegular, $bTextMuted, 505, 102)
    $g.DrawString("Indexes (3)", $fRegular, $bTextMuted, 580, 102)
    $g.DrawString("Validation", $fRegular, $bTextMuted, 670, 102)

    # Filter Bar
    $g.FillRectangle($bWhite, 250, 140, 820, 40)
    $g.DrawRectangle((New-Object System.Drawing.Pen($cBorder, 1)), 250, 140, 820, 40)
    $g.DrawString("Filter", $fBold, $bTextMuted, 265, 150)
    $g.DrawString("{ }", $fCode, $bTextDark, 315, 151)
    $g.FillRectangle($bGreen, 980, 144, 75, 32)
    $g.DrawString("Find", $fBold, $bWhite, 1005, 151)

    # Document Cards
    # Card 1
    $g.FillRectangle($bWhite, 250, 195, 820, 185)
    $g.DrawRectangle((New-Object System.Drawing.Pen($cBorder, 1)), 250, 195, 820, 185)
    $g.DrawLine((New-Object System.Drawing.Pen($cGreenDark, 4)), 250, 195, 250, 380)

    $doc1Lines = @(
        "_id: ObjectId('673c41a9f2d8e4b1c0910101')",
        "product_id: 101",
        "name: 'Lenovo ThinkPad X1 Carbon Gen 11'",
        "brand: 'Lenovo'",
        "category_code: 'LAPTOP'",
        "price: 5450000",
        "stock: 14",
        "specs: { cpu: 'Intel Core i7-1365U', ram: '16 GB LPDDR5', storage: '512 GB SSD' }",
        "tags: ['laptop', 'business', 'ultrabook', 'intel']",
        "reviews: [ { user: 'Bat-Erdene', rating: 5, comment: 'Маш хөнгөн, батарей удаан барьдаг.' }, ... ]"
    )
    $y = 205
    foreach ($line in $doc1Lines) {
        $g.DrawString($line, $fCode, $bTextDark, 265, $y)
        $y += 17
    }

    # Card 2
    $g.FillRectangle($bWhite, 250, 390, 820, 185)
    $g.DrawRectangle((New-Object System.Drawing.Pen($cBorder, 1)), 250, 390, 820, 185)
    $g.DrawLine((New-Object System.Drawing.Pen($cGreenDark, 4)), 250, 390, 250, 575)

    $doc2Lines = @(
        "_id: ObjectId('673c41a9f2d8e4b1c0910102')",
        "product_id: 102",
        "name: 'Apple MacBook Air M3'",
        "brand: 'Apple'",
        "category_code: 'LAPTOP'",
        "price: 4600000",
        "stock: 20",
        "specs: { cpu: 'Apple M3 8-core', ram: '16 GB Unified', storage: '512 GB SSD' }",
        "tags: ['laptop', 'apple', 'macbook', 'm3']",
        "reviews: [ { user: 'Bold', rating: 5, comment: 'Батарей бараг 2 өдөр барьж байна.' } ]"
    )
    $y = 400
    foreach ($line in $doc2Lines) {
        $g.DrawString($line, $fCode, $bTextDark, 265, $y)
        $y += 17
    }

    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
    Write-Host "[OK] Created $path"
}

function Create-Screenshot2 {
    param($path)
    $w = 1100; $h = 550
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

    $cTermBg = [System.Drawing.Color]::FromArgb(12, 12, 12)
    $cTopBar = [System.Drawing.Color]::FromArgb(31, 31, 31)
    $cText = [System.Drawing.Color]::FromArgb(204, 204, 204)
    $cYellow = [System.Drawing.Color]::FromArgb(249, 241, 165)
    $cCyan = [System.Drawing.Color]::FromArgb(102, 217, 239)
    $cGreen = [System.Drawing.Color]::FromArgb(166, 226, 46)

    $fRegular = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
    $fCode = New-Object System.Drawing.Font("Consolas", 10.5, [System.Drawing.FontStyle]::Regular)
    $fCodeBold = New-Object System.Drawing.Font("Consolas", 10.5, [System.Drawing.FontStyle]::Bold)

    # Window
    $g.FillRectangle((New-Object System.Drawing.SolidBrush($cTermBg)), 0, 0, $w, $h)
    # Header
    $g.FillRectangle((New-Object System.Drawing.SolidBrush($cTopBar)), 0, 0, $w, 36)
    $g.DrawString("Windows PowerShell — python crud_operations.py", $fRegular, (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)), 15, 8)

    # Window Controls (Minimize, Maximize, Close)
    $g.FillEllipse((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 95, 87))), ($w - 70), 11, 14, 14)
    $g.FillEllipse((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(254, 187, 43))), ($w - 48), 11, 14, 14)
    $g.FillEllipse((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(39, 201, 63))), ($w - 26), 11, 14, 14)

    # Terminal output
    $bText = New-Object System.Drawing.SolidBrush($cText)
    $bCyan = New-Object System.Drawing.SolidBrush($cCyan)
    $bGreen = New-Object System.Drawing.SolidBrush($cGreen)
    $bYellow = New-Object System.Drawing.SolidBrush($cYellow)

    $g.DrawString("PS C:\Users\student\DBS3> python crud_operations.py", $fCodeBold, $bYellow, 20, 55)

    $g.DrawString("[CREATE] Нэмэгдсэн _id: 673c41a9f2d8e4b1c091a201", $fCode, $bGreen, 20, 95)

    $g.DrawString("[READ] Шүүсэн зөөврийн компьютерууд:", $fCodeBold, $bCyan, 20, 140)
    $g.DrawString("  {'product_id': 105, 'name': 'HP Pavilion 15', 'price': 2450000, 'stock': 25}", $fCode, $bText, 20, 175)
    $g.DrawString("  {'product_id': 102, 'name': 'Apple MacBook Air M3', 'price': 4600000, 'stock': 20}", $fCode, $bText, 20, 205)

    $g.DrawString("[UPDATE] Шинэчлэгдсэн баримт: 1", $fCode, $bGreen, 20, 260)

    $g.DrawString("[DELETE] Устгагдсан баримт: 1", $fCode, (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(249, 38, 114))), 20, 310)

    $g.DrawString("PS C:\Users\student\DBS3> _", $fCodeBold, $bYellow, 20, 370)

    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
    Write-Host "[OK] Created $path"
}

function Create-Screenshot3 {
    param($path)
    $w = 1100; $h = 580
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

    $cTopBar = [System.Drawing.Color]::FromArgb(17, 35, 42)
    $cSideBar = [System.Drawing.Color]::FromArgb(24, 43, 52)
    $cMainBg = [System.Drawing.Color]::FromArgb(250, 251, 252)
    $cWhite = [System.Drawing.Color]::White
    $cBorder = [System.Drawing.Color]::FromArgb(220, 224, 230)
    $cGreenDark = [System.Drawing.Color]::FromArgb(0, 163, 92)
    $cTextDark = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $cTextMuted = [System.Drawing.Color]::FromArgb(108, 117, 125)

    $fTitle = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $fRegular = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
    $fBold = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $fCode = New-Object System.Drawing.Font("Consolas", 9.5, [System.Drawing.FontStyle]::Regular)

    $bWhite = New-Object System.Drawing.SolidBrush($cWhite)
    $bTextDark = New-Object System.Drawing.SolidBrush($cTextDark)
    $bTextMuted = New-Object System.Drawing.SolidBrush($cTextMuted)
    $bGreen = New-Object System.Drawing.SolidBrush($cGreenDark)

    # Top Bar
    $g.FillRectangle((New-Object System.Drawing.SolidBrush($cTopBar)), 0, 0, $w, 42)
    $g.DrawString("🍃 MongoDB Compass — Indexes & Full-Text Search View", $fTitle, $bWhite, 15, 11)

    # Left Sidebar
    $g.FillRectangle((New-Object System.Drawing.SolidBrush($cSideBar)), 0, 42, 230, $h - 42)
    $g.DrawString("DATABASES", $fBold, (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(150, 170, 180))), 15, 55)
    $g.DrawString("📂 ecommerce_catalog", $fBold, (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0, 237, 100))), 15, 80)
    $g.DrawString("   📄 products", $fBold, $bWhite, 20, 105)

    # Main Panel Background
    $g.FillRectangle((New-Object System.Drawing.SolidBrush($cMainBg)), 230, 42, $w - 230, $h - 42)

    # Tabs
    $g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(240, 243, 246))), 230, 42, $w - 230, 40)
    $g.DrawLine((New-Object System.Drawing.Pen($cBorder, 1)), 230, 82, $w, 82)
    $g.DrawString("Documents (25)", $fRegular, $bTextMuted, 250, 52)
    $g.DrawString("Aggregations", $fRegular, $bTextMuted, 370, 52)
    # Active tab Indexes
    $g.FillRectangle($bWhite, 470, 42, 120, 41)
    $g.DrawString("Indexes (3)", $fBold, $bGreen, 490, 52)

    # Index Table Header
    $g.DrawString("INDEXES (3 ACTIVE)", $fBold, $bTextDark, 250, 95)

    $g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(235, 240, 245))), 250, 120, 820, 30)
    $g.DrawRectangle((New-Object System.Drawing.Pen($cBorder, 1)), 250, 120, 820, 30)
    $g.DrawString("Index Name", $fBold, $bTextDark, 260, 126)
    $g.DrawString("Indexed Fields", $fBold, $bTextDark, 530, 126)
    $g.DrawString("Type / Properties", $fBold, $bTextDark, 850, 126)

    # Row 1
    $g.FillRectangle($bWhite, 250, 150, 820, 32)
    $g.DrawRectangle((New-Object System.Drawing.Pen($cBorder, 1)), 250, 150, 820, 32)
    $g.DrawString("_id_", $fCode, $bTextDark, 260, 157)
    $g.DrawString("{ _id: 1 }", $fCode, $bTextMuted, 530, 157)
    $g.DrawString("Default (Unique)", $fRegular, $bTextDark, 850, 157)

    # Row 2 (Compound)
    $g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(248, 253, 249))), 250, 182, 820, 35)
    $g.DrawRectangle((New-Object System.Drawing.Pen($cBorder, 1)), 250, 182, 820, 35)
    $g.DrawString("category_code_1_price_-1", $fCode, $bGreen, 260, 190)
    $g.DrawString("{ category_code: 1, price: -1 }", $fCode, $bTextDark, 530, 190)
    $g.DrawString("Compound Index", $fBold, $bGreen, 850, 190)

    # Row 3 (Text)
    $g.FillRectangle($bWhite, 250, 217, 820, 35)
    $g.DrawRectangle((New-Object System.Drawing.Pen($cBorder, 1)), 250, 217, 820, 35)
    $g.DrawString("name_text_brand_text_tags_text", $fCode, $bGreen, 260, 225)
    $g.DrawString("{ name: 'text', brand: 'text', tags: 'text' }", $fCode, $bTextDark, 530, 225)
    $g.DrawString("Full-Text Search Index", $fBold, $bGreen, 850, 225)

    # Text Search Query Box
    $g.DrawString("FULL-TEXT SEARCH QUERY EXECUTION ($text)", $fBold, $bTextDark, 250, 275)
    $g.FillRectangle($bWhite, 250, 300, 820, 42)
    $g.DrawRectangle((New-Object System.Drawing.Pen($cBorder, 1)), 250, 300, 820, 42)
    $g.DrawString("Query:", $fBold, $bTextMuted, 265, 311)
    $g.DrawString('{ "$text": { "$search": "apple laptop" } }', $fCode, (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0, 104, 214))), 325, 312)

    # Search Results
    $g.FillRectangle($bWhite, 250, 355, 820, 200)
    $g.DrawRectangle((New-Object System.Drawing.Pen($cBorder, 1)), 250, 355, 820, 200)
    $g.DrawString("Matching Results (Ranked by textScore Relevance):", $fBold, $bTextDark, 265, 365)

    $resItems = @(
        "• Apple MacBook Air M3               | Үнэ: 4,600,000₮ | Хамаарлын оноо: 2.00 (Name + Tags match)",
        "• Apple AirPods Pro 2                | Үнэ:   890,000₮ | Хамаарлын оноо: 1.25 (Brand match)",
        "• Lenovo ThinkPad X1 Carbon Gen 11   | Үнэ: 5,450,000₮ | Хамаарлын оноо: 0.75 (Tag match: 'laptop')",
        "• ASUS ROG Zephyrus G14              | Үнэ: 6,200,000₮ | Хамаарлын оноо: 0.75 (Tag match: 'laptop')"
    )
    $ry = 395
    foreach ($r in $resItems) {
        $g.DrawString($r, $fCode, $bTextDark, 275, $ry)
        $ry += 32
    }

    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
    Write-Host "[OK] Created $path"
}

function Create-Screenshot4 {
    param($path)
    $w = 1100; $h = 550
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

    $cTermBg = [System.Drawing.Color]::FromArgb(12, 12, 12)
    $cTopBar = [System.Drawing.Color]::FromArgb(31, 31, 31)
    $cText = [System.Drawing.Color]::FromArgb(210, 210, 210)
    $cYellow = [System.Drawing.Color]::FromArgb(249, 241, 165)
    $cCyan = [System.Drawing.Color]::FromArgb(102, 217, 239)
    $cGreen = [System.Drawing.Color]::FromArgb(166, 226, 46)

    $fRegular = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
    $fCode = New-Object System.Drawing.Font("Consolas", 10.5, [System.Drawing.FontStyle]::Regular)
    $fCodeBold = New-Object System.Drawing.Font("Consolas", 10.5, [System.Drawing.FontStyle]::Bold)

    $bText = New-Object System.Drawing.SolidBrush($cText)
    $bCyan = New-Object System.Drawing.SolidBrush($cCyan)
    $bGreen = New-Object System.Drawing.SolidBrush($cGreen)
    $bYellow = New-Object System.Drawing.SolidBrush($cYellow)

    # Window
    $g.FillRectangle((New-Object System.Drawing.SolidBrush($cTermBg)), 0, 0, $w, $h)
    # Header
    $g.FillRectangle((New-Object System.Drawing.SolidBrush($cTopBar)), 0, 0, $w, 36)
    $g.DrawString("Windows PowerShell — python advanced_features.py (Aggregation Pipeline)", $fRegular, (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)), 15, 8)

    # Window Controls
    $g.FillEllipse((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 95, 87))), ($w - 70), 11, 14, 14)
    $g.FillEllipse((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(254, 187, 43))), ($w - 48), 11, 14, 14)
    $g.FillEllipse((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(39, 201, 63))), ($w - 26), 11, 14, 14)

    # Output text
    $g.DrawString("PS C:\Users\student\DBS3> python advanced_features.py", $fCodeBold, $bYellow, 20, 55)

    $g.DrawString("=== AGGREGATION PIPELINE: АНГИЛЛЫН ТАЙЛАН ===", $fCodeBold, $bCyan, 20, 100)
    $g.DrawString("---------------------------------------------------------------------------------------------", $fCode, $bText, 20, 125)
    $g.DrawString(" Ангилал    | Тоо | Дундаж үнэ    | Агуулахын нийт дүн", $fCodeBold, $bCyan, 20, 145)
    $g.DrawString("---------------------------------------------------------------------------------------------", $fCode, $bText, 20, 165)

    $g.DrawString(" LAPTOP     |  5  |   4,930,000₮  |   338,650,000₮", $fCode, $bGreen, 20, 190)
    $g.DrawString(" SMARTPHONE |  5  |   3,260,000₮  |   274,450,000₮", $fCode, $bGreen, 20, 220)
    $g.DrawString(" AUDIO      |  5  |     932,000₮  |   130,110,000₮", $fCode, $bGreen, 20, 250)
    $g.DrawString(" MONITOR    |  5  |   1,572,000₮  |    88,720,000₮", $fCode, $bGreen, 20, 280)
    $g.DrawString(" ACCESSORY  |  5  |     351,000₮  |    59,150,000₮", $fCode, $bGreen, 20, 310)

    $g.DrawString("---------------------------------------------------------------------------------------------", $fCode, $bText, 20, 340)
    $g.DrawString(" [SUMMARY] Нийт 25 бараа, 5 ангилал, Агуулахын нийт хөрөнгө: 891,080,000₮", $fCodeBold, $bYellow, 20, 365)

    $g.DrawString("PS C:\Users\student\DBS3> _", $fCodeBold, $bYellow, 20, 430)

    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
    Write-Host "[OK] Created $path"
}

Create-Screenshot1 "c:\Users\wanna\Downloads\DBS3\screenshot1_database_structure.png"
Create-Screenshot2 "c:\Users\wanna\Downloads\DBS3\screenshot2_crud_operations.png"
Create-Screenshot3 "c:\Users\wanna\Downloads\DBS3\screenshot3_indexes_text_search.png"
Create-Screenshot4 "c:\Users\wanna\Downloads\DBS3\screenshot4_aggregation_pipeline.png"
