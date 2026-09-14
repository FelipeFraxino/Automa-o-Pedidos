import { createClient } from '@supabase/supabase-js'

const url = import.meta.env.VITE_SUPABASE_URL || 'https://cwxfxnbqhxnoxlnlgavt.supabase.co'
const key = import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY || 'sb_publishable_dMUPJobaqltkFEtOCw7Q5w_TXQrnC44'

export const supabase = createClient(url, key)
