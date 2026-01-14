import re
import xml.etree.ElementTree as ET
import unicodedata
from src.utils.latex_cleaner import normalize_input

NAMESPACES = {'m': 'http://www.w3.org/1998/Math/MathML'}
NS_URI = NAMESPACES['m']

def _register_namespaces():
    try:
        ET.register_namespace('', NS_URI)
    except:
        pass

_register_namespaces()

def _normalize_mathml_output(s: str) -> str:
    # Fix invalid XML entities (specifically unescaped &)
    # latex2mathml might output <mi>&</mi> for alignment tabs
    # Regex to find & not followed by entity pattern
    s = re.sub(r'&(?!(?:[a-zA-Z0-9]+|#[0-9]+|#x[0-9a-fA-F]+);)', '&amp;', s)
    
    # Replace \| with ‖ (U+2016) in mo elements
    # latex2mathml outputs \| for \Bigl\| but &#x02016; for \|
    # We want consistent ‖
    s = re.sub(r'(<mo[^>]*>)\s*\\\|\s*(?=</mo>)', r'\1‖', s)
    
    # Replace \{ and \} with { and } in mo elements
    s = re.sub(r'(<mo[^>]*>)\s*\\\{\s*(?=</mo>)', r'\1{', s)
    s = re.sub(r'(<mo[^>]*>)\s*\\\}\s*(?=</mo>)', r'\1}', s)
    
    try:
        # latex2mathml usually includes xmlns
        root = ET.fromstring(s)
    except ET.ParseError as e:
        # print(f"DEBUG: ParseError: {e}")
        # If parsing fails, return original string (or try to fix common issues)
        return s

    _strip_sized_fence_limits_for_word(root)

    # Transform the tree recursively
    _transform_element(root)

    if root.tag == f'{{{NS_URI}}}math' and len(root) == 1 and root[0].tag == f'{{{NS_URI}}}mrow' and not root[0].attrib:
        wrapper = root[0]
        root[:] = list(wrapper)

    # Ensure display="block"
    if 'display' not in root.attrib:
        root.set('display', 'block')
    else:
        root.set('display', 'block')

    return ET.tostring(root, encoding='unicode', short_empty_elements=True)

def _strip_sized_fence_limits_for_word(element):
    mo_tag = f'{{{NS_URI}}}mo'
    fence_chars = {'(', ')', '[', ']', '{', '}', '|', '‖', '⟨', '⟩', '<', '>'}
    for mo in element.iter(mo_tag):
        if 'minsize' not in mo.attrib and 'maxsize' not in mo.attrib:
            continue
        if (mo.text or '') not in fence_chars:
            continue
        mo.attrib.pop('minsize', None)
        mo.attrib.pop('maxsize', None)
        mo.set('fence', 'true')
        mo.set('stretchy', 'true')

def _normalize_unary_minus_for_word(element):
    children = list(element)
    new_children = []

    def is_unary_position(prev_node) -> bool:
        if prev_node is None:
            return True

        if prev_node.tag == f'{{{NS_URI}}}maligngroup':
            return True

        mi_tag = f'{{{NS_URI}}}mi'
        if prev_node.tag == mi_tag and not (prev_node.text or '').strip() and len(prev_node) == 0:
            return True

        mo_tag = f'{{{NS_URI}}}mo'
        if prev_node.tag == mo_tag:
            prev_text = prev_node.text or ''
            prev_form = prev_node.get('form')
            prev_fence = prev_node.get('fence')
            if prev_form == 'prefix':
                return True
            if prev_fence == 'true' and prev_form != 'postfix':
                return True
            return prev_text in ('(', '[', '{', ',', '=', '+', '−', '×', '·', '/', '*', ':', ';')

        if prev_node.tag in (f'{{{NS_URI}}}msup', f'{{{NS_URI}}}msub', f'{{{NS_URI}}}msubsup') and len(prev_node) > 0:
            base = prev_node[0]
            if base.tag == mo_tag:
                base_text = base.text or ''
                base_form = base.get('form')
                base_fence = base.get('fence')
                if base_form == 'prefix':
                    return True
                if base_fence == 'true' and base_form != 'postfix':
                    return True
                return base_text in ('(', '[', '{', ',', '=', '+', '−', '×', '·', '/', '*', ':', ';')

        return False

    i = 0
    while i < len(children):
        node = children[i]
        next_node = children[i + 1] if i + 1 < len(children) else None
        prev_node = new_children[-1] if new_children else None

        if node.tag == f'{{{NS_URI}}}mo':
            txt = node.text or ''
            if txt:
                open_fences = {'(', '[', '{', '⟨', '<'}
                close_fences = {')', ']', '}', '⟩', '>'}
                symmetric_fences = {'|', '‖'}
                is_fence = (
                    node.get('fence') == 'true'
                    or 'minsize' in node.attrib
                    or 'maxsize' in node.attrib
                    or txt in symmetric_fences
                )

                if is_fence:
                    if 'minsize' in node.attrib or 'maxsize' in node.attrib:
                        txt_is_symmetric = txt in symmetric_fences
                        if not txt_is_symmetric and 'minsize' in node.attrib and 'maxsize' in node.attrib:
                            if node.attrib.get('minsize') == node.attrib.get('maxsize'):
                                node.attrib.pop('maxsize', None)

                        node.set('fence', 'true')
                        node.set('stretchy', 'true')
                        if txt in symmetric_fences:
                            node.set('symmetric', 'true')

                        node.attrib.pop('form', None)
                        node.attrib.pop('lspace', None)
                        node.attrib.pop('rspace', None)

                        if element.tag != f'{{{NS_URI}}}mrow' or element.get('data-mjx-texclass') is None:
                            texclass = None
                            if txt in open_fences:
                                texclass = 'OPEN'
                            elif txt in close_fences:
                                texclass = 'CLOSE'
                            elif txt in symmetric_fences:
                                if element.tag in (f'{{{NS_URI}}}msub', f'{{{NS_URI}}}msubsup') and i == 0:
                                    texclass = 'CLOSE'
                                else:
                                    texclass = 'OPEN' if is_unary_position(prev_node) else 'CLOSE'

                            if texclass is not None:
                                wrapper = ET.Element(f'{{{NS_URI}}}mrow')
                                wrapper.set('data-mjx-texclass', texclass)
                                wrapper.append(node)
                                node = wrapper

        if node.tag == f'{{{NS_URI}}}mo' and (node.text or '') == '−' and next_node is not None:
            if is_unary_position(prev_node):
                if next_node.tag == f'{{{NS_URI}}}mn' and (next_node.text or '') and not (next_node.text or '').startswith(('−', '-')):
                    next_node.text = '−' + next_node.text
                    new_children.append(next_node)
                    i += 2
                    continue

                node.set('form', 'prefix')
                node.set('lspace', '0')
                node.set('rspace', '0')

        new_children.append(node)
        i += 1

    element[:] = new_children

def _normalize_bold_math_alphanum(element):
    mi_tag = f'{{{NS_URI}}}mi'
    mrow_tag = f'{{{NS_URI}}}mrow'

    children = list(element)
    changed = False
    for idx, child in enumerate(children):
        if child.tag != mi_tag:
            continue
        if child.get('mathvariant') is not None:
            continue
        text = child.text or ''
        if len(text) != 1:
            continue

        base = unicodedata.normalize('NFKD', text)
        if len(base) != 1 or base == text:
            continue

        name = unicodedata.name(text, '')
        if 'MATHEMATICAL' not in name:
            continue

        child.text = base
        child.set('mathvariant', 'bold')

        wrapper = ET.Element(mrow_tag)
        wrapper.set('data-mjx-texclass', 'ORD')
        wrapper.append(child)
        children[idx] = wrapper
        changed = True

    if changed:
        element[:] = children

def _normalize_mtable_layout(element):
    mtable_tag = f'{{{NS_URI}}}mtable'
    mtr_tag = f'{{{NS_URI}}}mtr'
    mtd_tag = f'{{{NS_URI}}}mtd'
    mrow_tag = f'{{{NS_URI}}}mrow'
    maligngroup_tag = f'{{{NS_URI}}}maligngroup'
    mo_tag = f'{{{NS_URI}}}mo'
    mtext_tag = f'{{{NS_URI}}}mtext'
    mi_tag = f'{{{NS_URI}}}mi'

    for mtable in element.iter(mtable_tag):
        rows = mtable.findall(mtr_tag)
        if not rows:
            continue

        structured = True
        for r in rows:
            mtds = r.findall(mtd_tag)
            if len(mtds) != 1:
                structured = False
                break
            inner = mtds[0].find(mrow_tag)
            if inner is None:
                structured = False
                break
            blocks = list(inner)
            if not blocks:
                structured = False
                break
            for b in blocks:
                if b.tag != mrow_tag or b.find(maligngroup_tag) is None:
                    structured = False
                    break
            if not structured:
                break

        if not structured:
            continue

        max_blocks = 0
        for r in rows:
            inner = r.find(mtd_tag).find(mrow_tag)
            if inner is None:
                continue

            blocks = list(inner)
            new_blocks = []
            prev_empty = False
            for blk in blocks:
                is_empty = blk.tag == mrow_tag and len(blk) == 1 and blk[0].tag == maligngroup_tag
                if is_empty and prev_empty:
                    continue
                new_blocks.append(blk)
                prev_empty = is_empty

            if len(new_blocks) != len(blocks):
                inner[:] = new_blocks

            if len(new_blocks) > max_blocks:
                max_blocks = len(new_blocks)

        if max_blocks < 2:
            continue

        for r in rows:
            inner = r.find(mtd_tag).find(mrow_tag)
            if inner is None:
                continue
            for blk in list(inner):
                if blk.tag != mrow_tag or len(blk) < 2:
                    continue
                if blk[0].tag != maligngroup_tag:
                    continue
                maybe_open = blk[1]
                if maybe_open.tag != mrow_tag or maybe_open.get('data-mjx-texclass') != 'OPEN':
                    continue
                has_norm = any((mo.text or '') in ('|', '‖') for mo in maybe_open.iter(mo_tag))
                if not has_norm:
                    continue
                if len(blk) >= 2 and blk[1].tag == mi_tag:
                    continue
                blk.insert(1, ET.Element(mi_tag))

        mtable.set('displaystyle', 'true')
        mtable.set(
            'columnalign',
            ' '.join('right' if idx % 2 == 0 else 'left' for idx in range(max_blocks)),
        )

        if max_blocks >= 4:
            mtable.set(
                'columnspacing',
                ' '.join('2em' if (i % 2 == 0) else '0em' for i in range(1, max_blocks)),
            )
        else:
            mtable.set('columnspacing', '0em')

        if len(rows) == 2 and max_blocks >= 4:
            has_sized = any(('minsize' in mo.attrib or 'maxsize' in mo.attrib) for mo in mtable.iter(mo_tag))
            if has_sized:
                mtable.set('rowspacing', '1.3em 0.3em')
            else:
                has_if = any((t.text or '').strip().startswith('if') for t in mtable.iter(mtext_tag))
                if has_if:
                    mtable.set('rowspacing', '0.9em 0.3em')

def _normalize_regular_mtable_layout(element):
    mtable_tag = f'{{{NS_URI}}}mtable'
    mtr_tag = f'{{{NS_URI}}}mtr'
    mtd_tag = f'{{{NS_URI}}}mtd'

    for mtable in element.iter(mtable_tag):
        rows = mtable.findall(mtr_tag)
        if not rows:
            continue

        has_multi_mtd = any(len(r.findall(mtd_tag)) != 1 for r in rows)
        if not has_multi_mtd:
            continue

        def is_empty_cell(mtd) -> bool:
            return len(mtd) == 0 and not (mtd.text or '').strip()

        cols_per_row = [r.findall(mtd_tag) for r in rows]
        if cols_per_row and all(len(cols) >= 2 for cols in cols_per_row):
            if all(is_empty_cell(cols[0]) for cols in cols_per_row):
                for r, cols in zip(rows, cols_per_row):
                    r.remove(cols[0])

                columnalign = (mtable.get('columnalign') or '').strip()
                if columnalign:
                    parts = columnalign.split()
                    parts = parts[1:] if len(parts) > 1 else ['left']
                    mtable.set('columnalign', ' '.join(parts))

                columnspacing = (mtable.get('columnspacing') or '').strip()
                if columnspacing:
                    parts = columnspacing.split()
                    parts = parts[1:] if len(parts) > 1 else []
                    if parts:
                        mtable.set('columnspacing', ' '.join(parts))
                    else:
                        mtable.attrib.pop('columnspacing', None)

        should_compress = False
        for r in rows:
            mtds = r.findall(mtd_tag)
            for a, b in zip(mtds, mtds[1:]):
                if is_empty_cell(a) and is_empty_cell(b):
                    should_compress = True
                    break
            if should_compress:
                break

        if not should_compress:
            continue

        for r in rows:
            mtds = r.findall(mtd_tag)
            new_mtds = []
            prev_empty = False
            for cell in mtds:
                empty = is_empty_cell(cell)
                if empty and prev_empty:
                    continue
                new_mtds.append(cell)
                prev_empty = empty
            if len(new_mtds) != len(mtds):
                r[:] = new_mtds

        max_cols = 0
        for r in rows:
            cols = len(r.findall(mtd_tag))
            if cols > max_cols:
                max_cols = cols

        if max_cols < 2:
            continue

        for r in rows:
            cols = r.findall(mtd_tag)
            while len(cols) < max_cols:
                r.append(ET.Element(mtd_tag))
                cols = r.findall(mtd_tag)

        mtable.set('displaystyle', 'true')
        mtable.set(
            'columnalign',
            ' '.join('right' if idx % 2 == 0 else 'left' for idx in range(max_cols)),
        )
        if max_cols >= 4:
            mtable.set(
                'columnspacing',
                ' '.join('2em' if (i % 2 == 0) else '0em' for i in range(1, max_cols)),
            )
        else:
            mtable.set('columnspacing', '0em')

def _normalize_ord_wrapper_for_bold(element):
    mrow_tag = f'{{{NS_URI}}}mrow'
    mi_tag = f'{{{NS_URI}}}mi'
    mn_tag = f'{{{NS_URI}}}mn'

    for mrow in element.iter(mrow_tag):
        if mrow.attrib:
            continue
        if len(mrow) != 1:
            continue
        child = mrow[0]
        if child.tag not in (mi_tag, mn_tag):
            continue
        if child.get('mathvariant') != 'bold':
            continue
        mrow.set('data-mjx-texclass', 'ORD')

def _normalize_nested_mtables(element):
    mtable_tag = f'{{{NS_URI}}}mtable'
    mtr_tag = f'{{{NS_URI}}}mtr'
    mtd_tag = f'{{{NS_URI}}}mtd'
    mrow_tag = f'{{{NS_URI}}}mrow'
    maligngroup_tag = f'{{{NS_URI}}}maligngroup'

    def convert_structured(mtable):
        rows = mtable.findall(mtr_tag)
        if not rows:
            return

        converted_rows = []
        for r in rows:
            mtds = r.findall(mtd_tag)
            if len(mtds) != 1:
                return
            inner = mtds[0].find(mrow_tag)
            if inner is None:
                return

            new_mtr = ET.Element(mtr_tag)
            for blk in list(inner):
                new_mtd = ET.Element(mtd_tag)
                if blk.tag == mrow_tag and len(blk) > 0 and blk[0].tag == maligngroup_tag:
                    content = list(blk)[1:]
                else:
                    content = list(blk)
                for node in content:
                    new_mtd.append(node)
                new_mtr.append(new_mtd)
            converted_rows.append(new_mtr)

        mtable[:] = converted_rows

    def walk(node, inside_mtable: bool):
        is_table = node.tag == mtable_tag
        if is_table and inside_mtable:
            convert_structured(node)
        for child in list(node):
            walk(child, inside_mtable or is_table)

    walk(element, False)

def _normalize_norm_ord_mo(element):
    mo_tag = f'{{{NS_URI}}}mo'
    for mo in element.iter(mo_tag):
        if (mo.text or '') not in ('|', '‖'):
            continue
        if mo.get('fence') == 'false' and mo.get('stretchy') == 'false':
            mo.set('data-mjx-texclass', 'ORD')

def _normalize_infty_mi(element):
    mi_tag = f'{{{NS_URI}}}mi'
    mo_tag = f'{{{NS_URI}}}mo'
    for mi in element.iter(mi_tag):
        if (mi.text or '') == '∞' and mi.get('mathvariant') is None:
            mi.set('mathvariant', 'normal')

    for mo in element.iter(mo_tag):
        if (mo.text or '') == '∞':
            mo.tag = mi_tag
            mo.attrib.clear()
            mo.set('mathvariant', 'normal')

def _normalize_transpose_operator(element):
    mo_tag = f'{{{NS_URI}}}mo'
    mi_tag = f'{{{NS_URI}}}mi'
    for mo in element.iter(mo_tag):
        if (mo.text or '') == '⊤':
            mo.tag = mi_tag
            mo.attrib.clear()
            mo.set('mathvariant', 'normal')

def _prune_empty_mstyles(element):
    mstyle_tag = f'{{{NS_URI}}}mstyle'

    def walk(node):
        for child in list(node):
            walk(child)

        for child in list(node):
            if child.tag != mstyle_tag:
                continue
            if len(child) != 0:
                continue
            if (child.text or '').strip():
                continue
            node.remove(child)

    walk(element)

def _normalize_texclass_wrapper_nesting(element):
    mrow_tag = f'{{{NS_URI}}}mrow'

    def walk(node):
        i = 0
        while i < len(node):
            child = node[i]
            if (
                child.tag == mrow_tag
                and child.get('data-mjx-texclass') is not None
                and len(child) == 1
                and child[0].tag == mrow_tag
                and child[0].get('data-mjx-texclass') is not None
            ):
                inner = child[0]
                child.remove(inner)
                node.remove(child)
                node.insert(i, inner)
                continue

            walk(child)
            i += 1

    walk(element)

def _normalize_sized_fence_texclass(element):
    mrow_tag = f'{{{NS_URI}}}mrow'
    maligngroup_tag = f'{{{NS_URI}}}maligngroup'
    mi_tag = f'{{{NS_URI}}}mi'
    mo_tag = f'{{{NS_URI}}}mo'
    msub_tag = f'{{{NS_URI}}}msub'
    msubsup_tag = f'{{{NS_URI}}}msubsup'

    def is_unary_position(prev_node) -> bool:
        if prev_node is None:
            return True
        if prev_node.tag == maligngroup_tag:
            return True
        if prev_node.tag == mi_tag and not (prev_node.text or '').strip() and len(prev_node) == 0:
            return True
        if prev_node.tag == mo_tag:
            prev_text = prev_node.text or ''
            prev_form = prev_node.get('form')
            prev_fence = prev_node.get('fence')
            if prev_form == 'prefix':
                return True
            if prev_fence == 'true' and prev_form != 'postfix':
                return True
            return prev_text in ('(', '[', '{', ',', '=', '+', '−', '×', '·', '/', '*', ':', ';')
        if prev_node.tag in (f'{{{NS_URI}}}msup', msub_tag, msubsup_tag) and len(prev_node) > 0:
            base = prev_node[0]
            if base.tag == mo_tag:
                base_text = base.text or ''
                base_form = base.get('form')
                base_fence = base.get('fence')
                if base_form == 'prefix':
                    return True
                if base_fence == 'true' and base_form != 'postfix':
                    return True
                return base_text in ('(', '[', '{', ',', '=', '+', '−', '×', '·', '/', '*', ':', ';')
        return False

    def walk(node):
        children = list(node)
        for idx, child in enumerate(children):
            if (
                child.tag == mrow_tag
                and child.get('data-mjx-texclass') in ('OPEN', 'CLOSE')
                and len(child) == 1
                and child[0].tag == mo_tag
            ):
                mo = child[0]
                txt = mo.text or ''
                is_sized = 'minsize' in mo.attrib or 'maxsize' in mo.attrib or mo.get('symmetric') == 'true'
                if txt in ('|', '‖') and is_sized:
                    if node.tag in (msub_tag, msubsup_tag) and idx == 0:
                        child.set('data-mjx-texclass', 'CLOSE')
                    else:
                        prev = children[idx - 1] if idx > 0 else None
                        child.set('data-mjx-texclass', 'OPEN' if is_unary_position(prev) else 'CLOSE')

            walk(child)

    walk(element)

def _flatten_table_markers(element):
    # Only flatten into containers that support inferred mrow (variable number of children)
    allowed_containers = {
        f'{{{NS_URI}}}mrow',
        f'{{{NS_URI}}}mstyle',
        f'{{{NS_URI}}}merror',
        f'{{{NS_URI}}}mphantom',
        f'{{{NS_URI}}}mpadded',
        f'{{{NS_URI}}}mtd',
        f'{{{NS_URI}}}math',
        f'{{{NS_URI}}}semantics',
        'math', # In case root tag doesn't have namespace in tag name (it usually does though)
    }
    
    if element.tag not in allowed_containers:
        return

    i = 0
    while i < len(element):
        child = element[i]
        # Check if child is a wrapper that we should flatten
        if child.tag in (f'{{{NS_URI}}}mstyle', f'{{{NS_URI}}}mrow'):
            # Check if child contains & or newline
            has_table_markers = False
            for grand in child:
                if (grand.tag == f'{{{NS_URI}}}mi' and grand.text == '&') or \
                   (grand.tag == f'{{{NS_URI}}}mspace' and grand.get('linebreak') == 'newline'):
                    has_table_markers = True
                    break
            
            if has_table_markers:
                # Flatten
                new_nodes = []
                for grand in child:
                    if (grand.tag == f'{{{NS_URI}}}mi' and grand.text == '&') or \
                       (grand.tag == f'{{{NS_URI}}}mspace' and grand.get('linebreak') == 'newline'):
                        new_nodes.append(grand)
                    else:
                        # Wrap in clone of child (mstyle/mrow)
                        # Only if child has attributes? 
                        # mrow with no attributes is transparent, but we should probably preserve it if it was there?
                        # Actually, if we flatten mrow, we lose the mrow boundary.
                        # If the mrow had NO attributes, it was just grouping.
                        # If it had attributes, we should preserve them.
                        # mstyle ALWAYS implies attributes (or at least intention).
                        
                        if child.tag == f'{{{NS_URI}}}mrow' and not child.attrib:
                             new_nodes.append(grand)
                        else:
                             wrapper = ET.Element(child.tag, child.attrib)
                             wrapper.append(grand)
                             new_nodes.append(wrapper)
                
                # Replace child with new_nodes
                element[i:i+1] = new_nodes
                # Don't increment i, process the new nodes (in case they are also flattenable wrappers? 
                # No, we wrapped them in new wrappers which contain only 1 child (not a marker).
                # So they won't match has_table_markers.
                # However, new_nodes might contain the MARKERS themselves.
                # We should skip the markers.
                # But it's easier to just increment i by len(new_nodes) or handle index carefully.
                # Actually, if we don't increment i, we check the first new node.
                # If it's '&', it's not mstyle/mrow, so we skip it (i+=1).
                # If it's mstyle(wrapper), it has no markers, so we skip it.
                # So not incrementing i is safe and correct.
                continue
        i += 1

def _transform_element(element):
    # Flatten mstyle/mrow containing table markers
    _flatten_table_markers(element)

    # Process children to find and replace fence pairs with mfenced
    new_children = []
    i = 0
    children = list(element)
    
    # Map of opening fence to closing fence
    # Note: latex2mathml produces entities or chars.
    # We should match what latex2mathml produces.
    # { -> } or &#x0007B; -> &#x0007D;
    # ( -> ) or &#x00028; -> &#x00029;
    # [ -> ] or &#x0005B; -> &#x0005D;
    # | -> | or &#x0007C; -> &#x0007C;
    # ⟨ -> ⟩ (angle brackets, usually &#x027E8; -> &#x027E9;)
    
    fence_pairs = {
        '{': '}', '\u007B': '\u007D',
        '(': ')', '\u0028': '\u0029',
        '[': ']', '\u005B': '\u005D',
        '|': '|', '\u007C': '\u007C',
        '‖': '‖', '\u2016': '\u2016',
        '⟨': '⟩', '\u27E8': '\u27E9',
        '<': '>', # Just in case
    }
    
    # Reverse map for checking closing chars
    # Note: multiple keys map to same value, but we need to know if a char IS a closing char
    # But strictly we match the pair started by opening char.
    
    while i < len(children):
        child = children[i]
        is_open_fence = False
        fence_char = ''
        
        # Check if child is an mo element
        # print(f"Processing child: {child.tag} text='{child.text}' attrib={child.attrib}") # DEBUG
        if child.tag == f'{{{NS_URI}}}mo':
            text = child.text
            # Check if it's a known opening fence
            if text in fence_pairs:
                form = child.get('form')
                fence_attr = child.get('fence')
                stretchy_attr = child.get('stretchy')
                
                # Only treat as open fence if:
                # 1. explicitly fence="true"
                # 2. OR stretchy is NOT "false" (implies stretchy or default)
                # 3. OR it's a known fence that latex2mathml might miss attributes on (like symmetric ones if any)
                
                should_be_fence = False
                if fence_attr == 'true':
                    should_be_fence = True
                elif stretchy_attr != 'false': 
                    # If stretchy is not false, it might be a fence.
                    # But latex2mathml might output stretchy="false" for small parens?
                    # No, inspect_l2m showed stretchy="false" for f(x).
                    should_be_fence = True
                
                # Treat as open fence unless explicitly marked as postfix
                # This covers latex2mathml output for bmatrix/vmatrix which has no attributes
                if should_be_fence and form != 'postfix':
                    # Skip if explicit sizing is present (e.g. \Bigl)
                    # We want to preserve minsize/maxsize which mfenced doesn't support well
                    if 'minsize' in child.attrib or 'maxsize' in child.attrib:
                        is_open_fence = False
                    else:
                        is_open_fence = True
                        fence_char = text
        
        if is_open_fence:
            # print(f"Found open fence: {fence_char}") # DEBUG
            # Search for matching closing fence
            j = i + 1
            balance = 1
            matched = False
            target_close = fence_pairs[fence_char]
            
            # For symmetric fences (like |), we need a heuristic to distinguish open/close
            # if attributes are missing.
            prev_was_open = True # children[i] was open
            
            # Variable to store the actual closing character found (could be empty for \right.)
            found_close_char = None
            
            while j < len(children):
                curr = children[j]
                
                curr_text = ''
                is_curr_open = False
                is_curr_close = False
                
                is_curr_open = False
                is_curr_close = False
                
                # Check for fence in curr node or wrapped in msup/msub/msubsup
                fence_node = None
                wrapped_type = None # None, 'sup', 'sub', 'subsup'
                wrapper_node = None
                
                if curr.tag == f'{{{NS_URI}}}mo':
                    fence_node = curr
                elif curr.tag in (f'{{{NS_URI}}}msup', f'{{{NS_URI}}}msub', f'{{{NS_URI}}}msubsup'):
                    # Check first child
                    if len(curr) > 0 and curr[0].tag == f'{{{NS_URI}}}mo':
                        fence_node = curr[0]
                        wrapper_node = curr
                        if curr.tag == f'{{{NS_URI}}}msup': wrapped_type = 'sup'
                        elif curr.tag == f'{{{NS_URI}}}msub': wrapped_type = 'sub'
                        elif curr.tag == f'{{{NS_URI}}}msubsup': wrapped_type = 'subsup'

                # New: Look inside mstyle/mrow for the closing fence if we haven't found it yet
                # This handles cases like \biggl( \displaystyle ... \biggr) where \biggr is inside mstyle
                match_inside_container = False
                container_match_idx = -1
                
                if fence_node is None and curr.tag in (f'{{{NS_URI}}}mstyle', f'{{{NS_URI}}}mrow'):
                     # Scan inside curr
                     sub_balance = balance
                     sub_prev_open = prev_was_open
                     
                     for k, sub in enumerate(curr):
                         sub_fence_node = None
                         if sub.tag == f'{{{NS_URI}}}mo':
                             sub_fence_node = sub
                         elif sub.tag in (f'{{{NS_URI}}}msup', f'{{{NS_URI}}}msub', f'{{{NS_URI}}}msubsup'):
                              if len(sub) > 0 and sub[0].tag == f'{{{NS_URI}}}mo':
                                  sub_fence_node = sub[0]
                         
                         if sub_fence_node is not None:
                             sub_text = sub_fence_node.text if sub_fence_node.text else ''
                             sub_form = sub_fence_node.get('form')
                             sub_fence = sub_fence_node.get('fence')
                             
                             is_sub_open = False
                             is_sub_close = False
                             
                             if not sub_text and (sub_form == 'postfix' or sub_fence == 'true'):
                                 is_sub_close = True
                             elif sub_text == fence_char:
                                 if sub_form == 'prefix' or (sub_fence == 'true' and sub_form != 'postfix'):
                                     is_sub_open = True
                                 elif not sub_form and not sub_fence:
                                     if fence_char == target_close:
                                         if sub_prev_open:
                                             is_sub_open = True
                                     else:
                                         is_sub_open = True
                             elif sub_text == target_close:
                                 if sub_form == 'postfix' or (sub_fence == 'true' and sub_form != 'prefix'):
                                      is_sub_close = True
                                 elif not sub_form and not sub_fence:
                                     if fence_char == target_close:
                                         if not sub_prev_open:
                                             is_sub_close = True
                                     else:
                                         is_sub_close = True
                             
                             if is_sub_open:
                                 sub_balance += 1
                                 sub_prev_open = True
                             if is_sub_close:
                                 sub_balance -= 1
                                 sub_prev_open = False
                             
                             if sub_balance == 0:
                                 # Found match inside curr at index k
                                 match_inside_container = True
                                 container_match_idx = k
                                 break
                     
                     # If we found it, we break the outer loop (simulated by setting fence_node to dummy)
                     # Actually we need to handle this specially.
                     if match_inside_container:
                         matched = True
                         # We need to construct the mfenced element here and now.
                         
                         # Split curr (mstyle/mrow)
                         # part1: curr children [:container_match_idx] -> inside mfenced
                         # part2: curr children [container_match_idx+1:] -> outside, after mfenced
                         
                         # Content of mfenced:
                         # 1. siblings between i+1 and j
                         # 2. part1 wrapped in new container (clone of curr)
                         
                         siblings_before = children[i+1:j]
                         
                         inner_part_1 = list(curr)[:container_match_idx]
                         
                         # Create clone of curr for part 1
                         # We must preserve attributes
                         container_clone_1 = ET.Element(curr.tag, curr.attrib)
                         container_clone_1.extend(inner_part_1)
                         
                         # Recursively process the inner content (part 1)
                         # Note: siblings_before is already processed? No.
                         # Wait, we are in _transform_element(element).
                         # children[i+1:j] are NOT processed yet (we are iterating i).
                         # So we need to process them.
                         
                         dummy1 = ET.Element('dummy')
                         dummy1.extend(siblings_before)
                         dummy1.append(container_clone_1)
                         _transform_element(dummy1)
                         processed_content = list(dummy1)
                         
                         # Create mfenced
                         mfenced = ET.Element(f'{{{NS_URI}}}mfenced')
                         mfenced.set('open', fence_char)
                         
                         # Determine close char
                         # The node at container_match_idx is the close fence
                         close_node_container = curr[container_match_idx]
                         # It might be wrapped
                         actual_close_node = close_node_container
                         if close_node_container.tag != f'{{{NS_URI}}}mo' and len(close_node_container) > 0:
                              actual_close_node = close_node_container[0]
                         
                         found_c_char = actual_close_node.text if actual_close_node.text else ''
                         # If it's empty fence
                         c_form = actual_close_node.get('form')
                         c_fence = actual_close_node.get('fence')
                         if not found_c_char and (c_form == 'postfix' or c_fence == 'true'):
                             found_c_char = ''
                         
                         close_val = found_c_char if found_c_char is not None else target_close
                         mfenced.set('close', close_val)
                         mfenced.set('separators', '|')
                         
                         # Wrap content in mrow
                         mrow = ET.Element(f'{{{NS_URI}}}mrow')
                         mrow.extend(processed_content)
                         mfenced.append(mrow)
                         
                         # Handle wrapper on close fence?
                         # If close_node_container was wrapped (msup etc), we need to wrap the mfenced?
                         # Yes.
                         final_node = mfenced
                         if close_node_container.tag != f'{{{NS_URI}}}mo':
                             new_wrapper = ET.Element(close_node_container.tag)
                             new_wrapper.append(mfenced)
                             for ch in list(close_node_container)[1:]:
                                 new_wrapper.append(ch)
                             final_node = new_wrapper
                         
                         new_children.append(final_node)
                         
                         # Now handle part 2 (remaining content of curr)
                         inner_part_2 = list(curr)[container_match_idx+1:]
                         if inner_part_2:
                             container_clone_2 = ET.Element(curr.tag, curr.attrib)
                             container_clone_2.extend(inner_part_2)
                             # We need to process this part too?
                             # Yes, recurse.
                             _transform_element(container_clone_2)
                             new_children.append(container_clone_2)
                         
                         i = j + 1
                         break
                     
                     # If not found inside, we treat curr as a normal block that affects balance
                     # sub_balance is the balance change caused by curr
                     # We update main balance
                     balance = sub_balance
                     prev_was_open = sub_prev_open

                if fence_node is not None:
                    curr_text = fence_node.text if fence_node.text else ''
                    form = fence_node.get('form')
                    fence = fence_node.get('fence')
                    
                    # Check for empty closing fence (matches anything, e.g. \right.)
                    if not curr_text and (form == 'postfix' or fence == 'true'):
                         is_curr_close = True
                         # If this balances out, we found our close
                         if balance == 1:
                             found_close_char = ''
                    
                    # Check for nested open
                    elif curr_text == fence_char:
                        if form == 'prefix' or (fence == 'true' and form != 'postfix'):
                            is_curr_open = True
                        elif not form and not fence:
                            # Heuristic for symmetric fence
                            if fence_char == target_close:
                                if prev_was_open:
                                    is_curr_open = True
                            else:
                                # Asymmetric, always open
                                is_curr_open = True

                    # Check for close
                    elif curr_text == target_close:
                        if form == 'postfix' or (fence == 'true' and form != 'prefix'):
                             is_curr_close = True
                        elif not form and not fence:
                            # Heuristic for symmetric fence
                            if fence_char == target_close:
                                if not prev_was_open:
                                    is_curr_close = True
                            else:
                                # Asymmetric, always close
                                is_curr_close = True

                # Apply balance changes
                if is_curr_open:
                    balance += 1
                
                if is_curr_close:
                    balance -= 1
                
                # Update heuristic state
                if is_curr_open:
                    prev_was_open = True
                elif is_curr_close:
                    prev_was_open = False # It was a close fence
                else:
                    # Content or unrelated char
                    prev_was_open = False
                
                if balance == 0:
                    matched = True
                    # Found match at index j
                    
                    # Recursively process the content between i+1 and j
                    inner_nodes = children[i+1:j]
                    # Create a temporary root to process inner nodes
                    dummy = ET.Element('dummy')
                    dummy.extend(inner_nodes)
                    _transform_element(dummy) # Recurse!
                    processed_inner = list(dummy)
                    
                    # Create mfenced
                    mfenced = ET.Element(f'{{{NS_URI}}}mfenced')
                    mfenced.set('open', fence_char)
                    # Use found_close_char if it was an empty fence, otherwise target_close
                    close_val = found_close_char if found_close_char is not None else target_close
                    mfenced.set('close', close_val)
                    mfenced.set('separators', '|') 
                    
                    # Check for newline to build mtable
                    # Note: We now do this generically at the end of _transform_element,
                    # but for mfenced content we might want to ensure it's wrapped in mrow if not a table.
                    # _build_table_if_needed is now called inside _transform_element.
                    # But processed_inner already has _transform_element called on it.
                    # So if it was a table, it is already a table.
                    
                    # We just wrap processed_inner.
                    # Unwrap mrow if it contains only mtable (latex2mathml often wraps aligned in mrow)
                    if len(processed_inner) == 1 and processed_inner[0].tag == f'{{{NS_URI}}}mrow':
                         if len(processed_inner[0]) == 1 and processed_inner[0][0].tag == f'{{{NS_URI}}}mtable':
                             processed_inner = [processed_inner[0][0]]

                    # Word prefers them in an mrow usually.
                    if len(processed_inner) == 1 and processed_inner[0].tag == f'{{{NS_URI}}}mtable':
                        mrow = ET.Element(f'{{{NS_URI}}}mrow')
                        mrow.append(processed_inner[0])
                        mfenced.append(mrow)
                    else:
                        mrow = ET.Element(f'{{{NS_URI}}}mrow')
                        mrow.extend(processed_inner)
                        mfenced.append(mrow)
                    
                    # Handle wrapper (msup/msub) on the closing fence
                    final_node = mfenced
                    if wrapped_type:
                        # Create new wrapper
                        new_wrapper = ET.Element(wrapper_node.tag)
                        # First child is the base (mfenced)
                        new_wrapper.append(mfenced)
                        # Copy other children (scripts)
                        # wrapper_node[0] is the fence mo, skip it
                        for child_node in list(wrapper_node)[1:]:
                             new_wrapper.append(child_node)
                        final_node = new_wrapper
                        
                    new_children.append(final_node)
                    i = j + 1
                    break
                
                j += 1
            
            if not matched:
                # Special handling for cases: { followed by mtable and no closing fence
                if fence_char == '{' and i + 1 < len(children) and children[i+1].tag == f'{{{NS_URI}}}mtable':
                     # Wrap { and mtable
                     # The mtable is children[i+1]. 
                     # We assume it is the content.
                     
                     # Recursively process the table (though it's likely already fine)
                     # But we should recurse to handle inner stuff
                     dummy = ET.Element('dummy')
                     dummy.append(children[i+1])
                     _transform_element(dummy)
                     processed_table = list(dummy)[0]

                     mfenced = ET.Element(f'{{{NS_URI}}}mfenced')
                     mfenced.set('open', '{')
                     mfenced.set('close', '')
                     mfenced.set('separators', '|')
                     
                     # Word seems to prefer mrow wrapping for aligned/cases
                     mrow = ET.Element(f'{{{NS_URI}}}mrow')
                     mrow.append(processed_table)
                     mfenced.append(mrow)
                     
                     new_children.append(mfenced)
                     i += 2
                else:
                    # Treat as normal child
                    _transform_element(child)
                    new_children.append(child)
                    i += 1
        else:
            # Not a fence, recurse
            _transform_element(child)
            new_children.append(child)
            i += 1
            
    # Update element children
    element[:] = new_children
    _normalize_unary_minus_for_word(element)
    _normalize_bold_math_alphanum(element)

    simplified = []
    for child in list(element):
        if child.tag == f'{{{NS_URI}}}mrow' and not child.attrib and len(child) == 1 and child[0].tag == f'{{{NS_URI}}}mfenced':
            simplified.append(child[0])
        else:
            simplified.append(child)
    element[:] = simplified

    _normalize_texclass_wrapper_nesting(element)
    _normalize_sized_fence_texclass(element)
    _normalize_mtable_layout(element)
    _normalize_regular_mtable_layout(element)
    _normalize_nested_mtables(element)
    _normalize_ord_wrapper_for_bold(element)
    _normalize_norm_ord_mo(element)
    _normalize_infty_mi(element)
    _normalize_transpose_operator(element)
    _prune_empty_mstyles(element)
    
    # Post-processing: Check if this element should become a table
    # Don't convert if it's already mtable or if it's the root math (unless necessary?)
    # Root math is usually block display.
    if element.tag != f'{{{NS_URI}}}mtable':
        table_node = _build_table_if_needed(element)
        if table_node is not None:
             element[:] = [table_node]

def _build_table_if_needed(element_or_nodes):
    # Support passing element or list of nodes
    if isinstance(element_or_nodes, list):
        nodes = element_or_nodes
    else:
        nodes = list(element_or_nodes)

    # If the content is wrapped in a single mrow, unwrap it to check for newlines
    target_nodes = nodes
    if len(nodes) == 1 and nodes[0].tag == f'{{{NS_URI}}}mrow':
        target_nodes = list(nodes[0])

    # Check for mspace linebreak="newline" or alignment tab &
    has_newline = False
    has_alignment = False
    
    for node in target_nodes:
        if node.tag == f'{{{NS_URI}}}mspace' and node.get('linebreak') == 'newline':
            has_newline = True
        if node.tag == f'{{{NS_URI}}}mi' and node.text == '&':
            has_alignment = True
            
    if not has_newline and not has_alignment:
        return None
        
    rows = []
    current_row = []
    
    for node in target_nodes:
        if node.tag == f'{{{NS_URI}}}mspace' and node.get('linebreak') == 'newline':
            rows.append(current_row)
            current_row = []
        else:
            current_row.append(node)
    rows.append(current_row)

    mtable = ET.Element(f'{{{NS_URI}}}mtable')
    if has_alignment:
        rows_cells = []
        for row_nodes in rows:
            cells = []
            current_cell = []
            for node in row_nodes:
                if node.tag == f'{{{NS_URI}}}mi' and node.text == '&':
                    cells.append(current_cell)
                    current_cell = []
                else:
                    current_cell.append(node)
            cells.append(current_cell)
            rows_cells.append(cells)

        max_cols = 0
        for cells in rows_cells:
            if len(cells) > max_cols:
                max_cols = len(cells)

        mtable.set('displaystyle', 'true')
        mtable.set(
            'columnalign',
            ' '.join('right' if idx % 2 == 0 else 'left' for idx in range(max_cols or 2)),
        )
        if (max_cols or 2) >= 4:
            mtable.set(
                'columnspacing',
                ' '.join('2em' if (i % 2 == 0) else '0em' for i in range(1, (max_cols or 2))),
            )
        else:
            mtable.set('columnspacing', '0em')
        mtable.set('rowspacing', '3pt')

        for cells in rows_cells:
            mtr = ET.Element(f'{{{NS_URI}}}mtr')
            row_mrow = ET.Element(f'{{{NS_URI}}}mrow')
            for cell_nodes in cells:
                cell_block = ET.Element(f'{{{NS_URI}}}mrow')
                cell_block.append(ET.Element(f'{{{NS_URI}}}maligngroup'))

                items = cell_nodes
                if len(items) == 1 and items[0].tag == f'{{{NS_URI}}}mrow':
                    items = list(items[0])
                for ch in items:
                    cell_block.append(ch)

                row_mrow.append(cell_block)

            mtd = ET.Element(f'{{{NS_URI}}}mtd')
            mtd.append(row_mrow)
            mtr.append(mtd)
            mtable.append(mtr)

        return mtable

    mtable.set('columnspacing', '1em')
    mtable.set('rowspacing', '4pt')

    for row_nodes in rows:
        mtr = ET.Element(f'{{{NS_URI}}}mtr')
        mtd = ET.Element(f'{{{NS_URI}}}mtd')
        if row_nodes:
            mtd.extend(row_nodes)
        mtr.append(mtd)
        mtable.append(mtr)

    return mtable


def convert(latex: str) -> str:
    cleaned = normalize_input(latex)
    from latex2mathml.converter import convert as l2m_convert
    return _normalize_mathml_output(l2m_convert(cleaned))
