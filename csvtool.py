#!/usr/bin/env python

# Copyright (c) 2004-2006 Tyler C. Sarna
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions
# are met:
#
# 1. Redistributions of source code must retain the above copyright
#    notice, this list of conditions and the following disclaimer.
# 2. Redistributions in binary form must reproduce the above copyright
#    notice, this list of conditions and the following disclaimer in the
#    documentation and/or other materials provided with the distribution.
# 3. All advertising materials mentioning features or use of this software
#    must display the following acknowledgement:
#        This product includes software developed by Tyler C. Sarna
#        http://ty.sarna.org/
# 4. Neither the name of Tyler C. Sarna nor the names of
#    contributors may be used to endorse or promote products derived
#    from this software without specific prior written permission. 
#    
# THIS SOFTWARE IS PROVIDED BY TYLER C. SARNA AND CONTRIBUTORS
# ``AS IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED
# TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
# PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL TYLER C. SARNA OR CONTRIBUTORS
# BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
# CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
# SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
# INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
# ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
# POSSIBILITY OF SUCH DAMAGE.

import sys, csv
from optparse import OptionParser

######## Reader/Writer with NULL handling

class Reader(object):
    def __init__(self, f, **kw):
        self.f = f

        self.null = kw.get('null')
        del kw['null']
        
        self.r = csv.reader(f, **kw)

    def __iter__(self):
        return self
        
    def next(self):
        row = self.r.next()
        l = []
        for c in row:
            if c == self.null:
                c = None
            l.append(c)             
        return l

    
class Writer(object):
    def __init__(self, f, **kw):
        self.f = f

        self.null = kw.get('null', '')
        print '--->', `self.null`
        del kw['null']
        
        self.w = csv.writer(f, **kw)

    def _fixrow(self, r):
        l = []
        for v in r:
            if v is None:
                v = self.null
            l.append(v)
        return l
        
    def writerow(self, r):
        return self.w.writerow(self._fixrow(r)) 
    
    def writerows(self, rl):
        l = [self._fixrow(r) for r in rl]
        self.w.writerows(l)
        
    
######## Command base class

class cmd(object):
    usage = "%prog"
    
    def __init__(self, prog, inargs):
        self.prog, self.inargs = prog, inargs
        self.prog_name = self.prog.prog_name + ' [global-options] ' + self.inargs[0]

        self.op = OptionParser(
            prog = self.prog_name,
            version = "%prog " + self.prog.version,
            usage = self.usage
        )

    def parse(self):
        self.opts, self.args = self.op.parse_args(self.inargs[1:])

    def parse_error(self, err):        
        self.op.print_usage(self.prog.stderr)
        self.prog.stderr.write("%s %s: error: %s\n" % (self.prog.prog_name, self.inargs[0], err))

        return 2

    def toStr(self, v):
        if v is None:
            v = self.prog.g_opts.output_null
        return str(v)
        
    def rowToStrs(self, r):
        return [self.toStr(v) for v in r]
        
    def parse_widths(self, s):
        s = s.split(',')
        w = []
        for i in s:
            if i:
                w.append(-int(i))
            else:
                w.append(0)

        self.widths = w

    def update_widths(self, r):
        if len(r) > len(self.widths):
            self.widths.extend([0] * (len(r) - len(self.widths)))
            
        for ix, v in enumerate(r):
            if self.widths[ix] < 0:
                pass # fixed width
            else:
                self.widths[ix] = min(max(self.widths[ix], len(v)), self.opts.max_width)

    def finalize_widths(self):
        self.widths = [abs(x) for x in self.widths]



######## cat

class cat_cmd(cmd):
    short_desc = "Catenate"
    usage = "%prog [options] [file [file [...]]]"

    def __call__(self):
        self.op.add_option("-r", "--remove-headers",
            action="store_true", default=False,
            help="Remove headers from second and later files")

        self.parse()
        
        cout = self.prog.writer()

        self.additional = False
        if not self.args:
            self.do_cat(self.prog.reader(), cout)
        else:
            for fn in self.args:
                self.do_cat(self.prog.reader(fn), cout)
                self.additional = True

        return 0

    def do_cat(self, cin, cout):
        if self.additional and self.opts.remove_headers:
            cin.next()
            
        for r in cin:
            cout.writerow(r)
                            

######## pivot

class pivot_cmd(cmd):
    short_desc = "Pivot table"
    usage = "%prog -x COLLIST -y COLLIST -z COLSPEC"
    
    def __call__(self):
        self.op.add_option("-x", "--columns",
            metavar="COLLIST",
            help="column list for output columns")
            
        self.op.add_option("-y", "--rows",
            metavar="COLLIST",
            help="column list for output rows")
            
        self.op.add_option("-z", "--value",
            metavar="COLSPEC",
            help="column output values")

        self.parse()

        if self.args:
            return self.parse_error("takes no arguments")

        if self.opts.value is None or (
            self.opts.rows is None and self.opts.columns is None):
            return self.parse_error("at least -z and one of -x or -y must be specified")
                    
        cin = self.prog.reader()
        incols = cin.next()

        columns = self.prog.parse_collist(incols, self.opts.columns)
        rows = self.prog.parse_collist(incols, self.opts.rows)
        value_col = self.prog.parse_colspec(incols, self.opts.value)

        res = {}
        seencols = {}
        
        for row in cin:
            x = tuple([row[c] for c in columns])
            y = tuple([row[c] for c in rows])
            seencols[x] = 1
            v = self.prog.to_numeric(row[value_col])
        
            yr = res.setdefault(y, {})
            yr[x] = yr.get(x, 0) + v
            
        outcols = seencols.keys()
        outcols.sort()

        outrows = res.keys()
        outrows.sort()

        cout = self.prog.writer()

        h = [incols[x] for x in rows]
        if not columns:
            h.append(incols[value_col])
        else:
            for c in outcols:
                h.append('-'.join(c))

        cout.writerow(h)

        for r in outrows:
            h = list(r)
            yr = res[r]
            for c in outcols:
                h.append(yr.get(c))

            cout.writerow(h)

######## sort

class sort_cmd(cmd):
    short_desc = "Sort table"
    usage = "%prog SORTSPEC"
    
    def __call__(self):
        self.parse()

        if len(self.args) != 1:
            return self.parse_error("sortspec required")
        
        cin = self.prog.reader()
        h = cin.next()

        self.parse_sortspec(h, self.args[0])

        t = []
        n = self.numeric
        tonum = self.prog.to_numeric
        for r in cin:
            for c in n:
                r[c] = tonum(r[c])
            t.append(r)

        for k, r in self.sorts:
            t.sort(key=(lambda x,k=k: x[k]), reverse=r)

        cout = self.prog.writer()

        cout.writerow(h)
        
        for r in t:
            cout.writerow(r)

    def parse_sortspec(self, incols, spec):
        spec = spec.split(',')

        self.numeric = []
        self.sorts = []
        
        for c in spec:
            n = r = False
            while c[0] in '-+#':
                if c[0] == '-':
                    r = True
                elif c[0] == '+':
                    r = False
                elif c[0] == '#':
                    n = True
                c = c[1:]
            c = self.prog.parse_colspec(incols, c)
            
            if n:
                self.numeric.append(c)
            self.sorts.append((c, r))

        self.sorts.reverse()

######## tocopy

class tocopy_cmd(cmd):
    short_desc = 'Convert to PostgreSQL/SQLite "COPY FROM" compatible format'
    usage = "%prog [options]"
    
    _map = {
        '\\' : '\\\\',  '\b' : '\\b',   '\f' : '\\f',
        '\n' : '\\n',   '\r' : '\\r',   '\t' : '\\t',   '\v' : '\\v'
    }
            
    def __call__(self):
        self.parse()    

        if self.args:
            return self.parse_error("takes no arguments")

        cin = self.prog.reader()
        out = self.prog.stdout.write
        
        delim = '\t'
        null = '\\N'

        h = cin.next() # discard header

        for r in cin:
            out('\t'.join(self.rowToStrs(r)))
            out('\n')
            
    def toStr(self, v):
        if v is None:
            return '\\N'

        v = str(v)
        nv = []
        for ch in v:
            o = ord(ch)
            m = self._map.get(ch)
            if m is not None:
                nv.append(m)
            elif o < 32 or o > 126:
                nv.append('\\%3o' % o)
            else:
                nv.append(ch)
        
        return ''.join(nv)
                    

######## tofancy

class tofancy_cmd(cmd):
    short_desc = "Format in fancy fashion"
    usage = "%prog [options]"
    
    def __call__(self):
        self.op.add_option("-m", "--max-width",
            metavar="WIDTH",
            type = "int",
            default = sys.maxint,
            help="Maximum column width")

        self.op.add_option("-w", "--widths",
            metavar="WIDTHS",
            help="fixed column widths separated by commas, leave a column blank to auto-size")

        self.parse()

        if self.args:
            return self.parse_error("takes no arguments")

        self.widths = []
        if self.opts.widths:
            self.parse_widths(self.opts.widths)
            
        cin = self.prog.reader()
        out = self.prog.stdout.write

        h = self.rowToStrs(cin.next())

        self.update_widths(h)

        t = []
        for r in cin:
            r = self.rowToStrs(r)
            self.update_widths(r)
            t.append(r)

        self.finalize_widths()

        out(self.single_sep())
        out(self.fmt_row(h))
        out(self.double_sep())
        
        for r in t:
            out(self.fmt_row(r))
            out(self.single_sep())

    def single_sep(self):
        return '+' + '+'.join(['-' * w for w in self.widths]) + '+\n'

    def double_sep(self):
        return '+' + '+'.join(['=' * w for w in self.widths]) + '+\n'

    def fmt_row(self, r):
        nr = []
        for ix, w in enumerate(self.widths):
            nr.append((r[ix] + ' ' * w)[:w])
        
        return '|' + '|'.join(nr) + '|\n'                
        
######## tohoriz

class tohoriz_cmd(cmd):
    short_desc = "Format in horizontal format"
    usage = "%prog [options]"
    
    def __call__(self):
        self.op.add_option("-m", "--max-width",
            metavar="WIDTH",
            type = "int",
            default = sys.maxint,
            help="Maximum column width")

        self.op.add_option("-w", "--widths",
            metavar="WIDTHS",
            help="fixed column widths separated by commas, leave a column blank to auto-size")

        self.parse()

        if self.args:
            return self.parse_error("takes no arguments")

        self.widths = []
        if self.opts.widths:
            self.parse_widths(self.opts.widths)
            
        cin = self.prog.reader()
        out = self.prog.stdout.write

        h = self.rowToStrs(cin.next())

        self.update_widths(h)

        t = [h, None]
        for r in cin:
            r = self.rowToStrs(r)
            self.update_widths(r)
            t.append(r)

        self.finalize_widths()

        t[1] = [('-' * w) for w in self.widths]
        
        for r in t:
            nr = []
            for ix, w in enumerate(self.widths):
                nr.append((r[ix] + ' ' * w)[:w])
                
            out(' '.join(nr))
            out('\n')

######## tohtml

class tohtml_cmd(cmd):
    short_desc = "Format in HTML"
    usage = "%prog [options]"
    
    def __call__(self):
        from cgi import escape
        
        self.op.add_option("-p", "--full-page",
            action="store_true", default=False,
            help="Generate complete html PAGE instead of fragment")

        self.op.add_option("-r", "--right-justify",
            metavar="COLLIST",
            help="column list to right-justify")
            
        self.op.add_option("-t", "--title",
            default="",
            metavar="TITLE",
            help="Add title to table")
            
        self.parse()

        if self.args:
            return self.parse_error("takes no arguments")

        cin = self.prog.reader()
        out = self.prog.stdout.write

        h = cin.next()
        self.rjust = self.prog.parse_collist(h, self.opts.right_justify)

        if self.opts.full_page:
            out("<html><head><title>%s</title></head><body>\n" % self.opts.title)

        out("<table border=1>\n")
        
        if self.opts.title:
            out("<tr><th colspan=%d>%s</th></tr>\n" % (len(h), escape(self.opts.title)))
        
        out(self.rowToHTML(escape, h, 'th'))
        
        for r in cin:
            out(self.rowToHTML(escape, r, 'td'))
            
        out("</table>\n")
        
        if self.opts.full_page:
            out("</body></html>\n")

    def rowToHTML(self, escape, r, tag):
        toStr = self.toStr
        nr = []
        for ix, c in enumerate(r):
            j = ''
            if ix in self.rjust:
                j = ' align="right"'
                
            nr.append('<%s%s>%s</%s>' % (tag, j, escape(toStr(c)), tag))

        return '<tr>%s</tr>\n' % ''.join(nr)

            
######## toinsert

class toinsert_cmd(cmd):
    short_desc = "Generate SQL INSERT statements"
    usage = "%prog [options] TABLENAME"
    
    def __call__(self):
        self.op.add_option("-n", "--noquote-columns",
            metavar="COLLIST",
            help="column list for numeric or other no-quotes columns")
            
        self.parse()

        if len(self.args) != 1:
            return self.parse_error("table name required")

        cin = self.prog.reader()
        out = self.prog.stdout.write

        h = cin.next()
        stmt = 'INSERT INTO %s (%s) VALUES (%%s)\ngo\n' % (
            self.args[0], ', '.join(h)
        )

        nqcols = self.prog.parse_collist(h, self.opts.noquote_columns)

        for r in cin:
            nr = []
            for ix, c in enumerate(r):
                if c is None:
                    nr.append('NULL')
                else:
                    c = str(c)
                    if ix not in nqcols:
                        c = '"%s"' % (c.replace('"', '""'))
                    nr.append(c)                

            out(stmt % (', '.join(nr)))

######## toldif

class toldif_cmd(cmd):
    short_desc = "Convert to LDIF"
    usage = "%prog"
    
    def __call__(self):
        self.parse()    

        if self.args:
            return self.parse_error("takes no arguments")

        cin = self.prog.reader()
        out = self.prog.stdout.write
        
        # dn must come first, according to RFC2849...
        cols = [(n, ix) for ix, n in enumerate(cin.next())]
        d = dict(cols)
        dn = 'dn'
        dnix = d.get(dn)
        if dnix is None:
            dn = 'DN'
            dnix = d.get(dn)
            
        if dnix:
            cols.remove((dn, dnix))
            cols.insert(0, (dn, dnix))
        else:
            pass # ...though we can't be truly compliant without one

        nr = 0
        for r in cin:
            nr += 1
            for n, ix in cols:
                v = r[ix]
                colname = '%s: ' % n
            
                if v is not None:
                    v = str(v)
                    ascii = True
                    for ch in v:
                        o = ord(ch)
                        if o < 32 or o > 126:
                            ascii = False
                            break
                    
                    if not ascii:
                        colname = '%s:: ' % n
                        v = ''.join(v.encode('base64').split())
                    
                    fl = 77 - len(colname)
                    out(colname + v[:fl] + '\n')
                    
                    v = v[fl:]
                    while v:
                        out(' ' + v[:76] + '\n')
                        v = v[76:]
            
            out('\n')

######## toupdate

class toupdate_cmd(cmd):
    short_desc = "Generate SQL UPDATE statements"
    usage = "%prog [options] TABLENAME KEYCOLLLIST"
    
    def __call__(self):
        self.op.add_option("-i", "--insert-or-update",
            action="store_true", default=False,
            help="Update if present else insert")

        self.op.add_option("-n", "--noquote-columns",
            metavar="COLLIST",
            help="column list for numeric or other no-quotes columns")
            
        self.parse()

        if len(self.args) != 2:
            return self.parse_error("table name and key column list required")

        cin = self.prog.reader()
        out = self.prog.stdout.write

        h = cin.next()

        nqcols = self.prog.parse_collist(h, self.opts.noquote_columns)
        kcols = self.prog.parse_collist(h, self.args[1])

        if not len(kcols):
            return self.parse_error("At least one key column required")

        update = 'UPDATE %s SET %%s\n\tWHERE %%s' % self.args[0] 
        insert = 'INSERT INTO %s (%%s)\n\tVALUES (%%s)' % self.args[0] 
        ioupdate = 'IF EXISTS (SELECT * FROM %s WHERE %%s)\n\t%s\nELSE %s' % (
            self.args[0], update, insert
        )

        iou = self.opts.insert_or_update
        ht = ', '.join(h)

        for r in cin:
            kl = []; ul = []; vl = []
            for ix, c in enumerate(r):
                if c is None:
                    c = 'NULL'
                else:
                    c = str(c)
                    if ix not in nqcols:
                        c = '"%s"' % (c.replace('"', '""'))

                vl.append(c)

                c = '%s=%s' % (h[ix], c)
                
                if ix in kcols:
                    kl.append(c)
                else:
                    ul.append(c)

            ul = ', '.join(ul)
            kl = ' AND '.join(kl)

            if iou:
                out(ioupdate % (kl, ul, kl, ht, ', '.join(vl)))
            else:
                out(update % (ul, kl))
                
            out('\ngo\n')

######## tovert

class tovert_cmd(cmd):
    short_desc = "Format in vertical style"
    usage = "%prog [options]"
    
    def __call__(self):
        self.op.add_option("-l", "--left-justify",
            action="store_true", default=False,
            help="Left justify column names")

        self.op.add_option("-m", "--max-width",
            metavar="WIDTH",
            type = "int",
            help="Maximum column name width")

        self.op.add_option("-s", "--separator",
            metavar="SEPARATOR",
            default=' ',
            help="separator between name and value (default '%default')")

        self.parse()    

        if self.args:
            return self.parse_error("takes no arguments")

        cin = self.prog.reader()
        out = self.prog.stdout.write
        
        h = cin.next()
        w = max([len(c) for c in h])
        if self.opts.max_width is not None:
            w = min(w, self.opts.max_width)

        l = self.opts.left_justify
        s = self.opts.separator
        
        first = True
        for r in cin:
            if not first:
                out('\n')
            else:
                first = False
                
            for ix, v in enumerate(r):
                c = h[ix][:w]
                if l:
                    c = c.ljust(w)
                else:
                    c = c.rjust(w)
                out("%s%s%s\n" % (c, s, self.toStr(v)))

######## main

class Main(object):
    version = '0.41'

    cmdlist = {
        'cat'       : cat_cmd,
        'pivot'     : pivot_cmd,
        'sort'      : sort_cmd,
        'tocopy'    : tocopy_cmd,
        'tofancy'   : tofancy_cmd, 
        'tohoriz'   : tohoriz_cmd, 
        'tohtml'    : tohtml_cmd, 
        'toinsert'  : toinsert_cmd, 
        'toldif'    : toldif_cmd,
        'toupdate'  : toupdate_cmd,
        'tovert'    : tovert_cmd,
    }

    def reader(self, filename=None):
        if filename is None or filename == '-':
            f = self.stdin
        else:
            f = open(filename, 'r')
        
        return Reader(f, null=self.g_opts.input_null, dialect=self.g_opts.input_dialect)

    def writer(self, filename=None):
        if filename is None or filename == '-':
            f = self.stdout
        else:
            f = open(filename, 'w')
        
        return Writer(f, null=self.g_opts.output_null, dialect=self.g_opts.output_dialect)

    def parse_colspec(self, avail_cols, colspec):
        # a number indicates the nth column, starting with 1 at the left
        # negative numbers count from the right
        # 0 is not valid
        i = None
        try:
            i = int(colspec)
        except ValueError:
            pass
        
        if i is None:
            # try exact match by name
            for i, c in enumerate(avail_cols):
                if c == colspec:
                    return i

            # try case-insensitive match
            for i, c in enumerate(avail_cols):
                if c.lower() == colspec.lower():
                    return i
            
            raise ValueError, "Cannot find column named '%s'" % colspec
        else:
            if i == 0:
                raise ValueError, "Column numbers start at 1"
            elif i > 0:
                i -= 1

        return i
                    
    def parse_collist(self, avail_cols, colspec):
        if not colspec:
            return []

        il = []
        for c in colspec.split(','):
            il.append(self.parse_colspec(avail_cols, c))
        
        return il

    def to_numeric(self, v):
        try:
            return int(v)
        except ValueError:
            return float(v)

    def __init__(self, args, stdin, stdout, stderr):
        self.args, self.stdin, self.stdout, self.stderr = \
            args, stdin, stdout, stderr

    def __call__(self):
        op = OptionParser(
            version = "%prog " + self.version,
            usage = "%prog [options] command ..."
        )

        op.disable_interspersed_args()
        
        op.add_option("-I", "--input-dialect",
            default="excel", metavar="DIALECT",
            help="CSV input dialect is DIALECT (default %default)")
            
        op.add_option("-O", "--output-dialect",
            default="excel", metavar="DIALECT",
            help="CSV output dialect is DIALECT (default %default)")
            
        op.add_option("-X", "--input-null",
            default=None, metavar="NULL-VALUE",
            help="Input representation of NULL (default %default)")
            
        op.add_option("-N", "--output-null",
            default="", metavar="NULL-VALUE",
            help="Output representation of NULL (default '%default')")
            
        self.g_opts, args = op.parse_args(self.args[1:])
        self.prog_name = op.get_prog_name()
    
        if len(args) < 1:
            op.print_usage(sys.stderr)
            sys.stderr.write("%s: error: command not specified\n" % self.prog_name)
            self.print_cmdlist()

            return 2

        cmdclass = self.cmdlist.get(args[0])
        if cmdclass is None:
            op.print_usage(sys.stderr)
            sys.stderr.write("%s: error: unknown command %s\n" % (op.get_prog_name(), args[0]))
            self.print_cmdlist()

            return 2

        cmd = cmdclass(self, args)
        
        return cmd()

    def print_cmdlist(self):
        sys.stderr.write("\nValid commands:\n")
        cl = self.cmdlist.items()
        cl.sort()
        for k, v in cl:
            sys.stderr.write("  %-16s%s\n" % (k, v.short_desc))
        
######## Run

if __name__ == '__main__':
    sys.exit(Main(sys.argv, sys.stdin, sys.stdout, sys.stderr)())
