export type RiskCategory =
  | 'FACT_BASED_ISSUES' | 'GOVERNMENT' | 'LIKENESS' | 'LOCATIONS'
  | 'MARKETING_ADDED_VALUE' | 'MUSIC_CHOREOGRAPHY' | 'NAMES_NUMBERS'
  | 'PLAYBACK' | 'PRODUCT_MISUSE' | 'PROPS_SET_DRESSING'
  | 'REFERENCES' | 'VEHICLES' | 'WARDROBE' | 'OTHER'

export type RiskSubCategory = string

export type RiskSeverity = 'HIGH' | 'MEDIUM' | 'LOW'

export type ClearanceStatus =
  | 'PENDING' | 'CLEARED' | 'NOT_CLEAR' | 'NEGOTIATED_BY_ATTORNEY'
  | 'BRANDED_INTEGRATION' | 'NO_CLEARANCE_NECESSARY' | 'PERMISSIBLE'

export type ScriptStatus = 'PROCESSING' | 'COMPLETE' | 'FAILED'

export interface RiskFlag {
  id: number
  category: RiskCategory
  subCategory: RiskSubCategory
  severity: RiskSeverity
  status: ClearanceStatus
  entityName: string
  snippet: string | null
  reason: string
  suggestion: string | null
  comments: string | null
  restrictions: string | null
  pageNumber: number
  episodeNumber: string | null
  sceneNumber: string | null
  isRedacted: boolean
  createdAt: string
}

export interface Script {
  id: number
  filename: string
  totalPages: number
  riskCount: number
  status: ScriptStatus
  uploadedAt: string
  risks: RiskFlag[] | null
}

export interface RiskUpdatePayload {
  status?: ClearanceStatus
  comments?: string
  restrictions?: string
  isRedacted?: boolean
}
