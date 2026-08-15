export type Json =
  | string
  | number
  | boolean
  | null
  | { [key: string]: Json | undefined }
  | Json[]

export type Database = {
  // Allows to automatically instantiate createClient with right options
  // instead of createClient<Database, { PostgrestVersion: 'XX' }>(URL, KEY)
  __InternalSupabase: {
    PostgrestVersion: "14.5"
  }
  public: {
    Tables: {
      admin_affairs_assets: {
        Row: {
          asset_location: string | null
          created_at: string
          id: string
          item_name: string | null
          purchase_order_no: string | null
          quantity_added: number | null
          request_date: string | null
          requesting_department: string | null
          supplier: string | null
          unit: string | null
          updated_at: string
          year: string | null
        }
        Insert: {
          asset_location?: string | null
          created_at?: string
          id?: string
          item_name?: string | null
          purchase_order_no?: string | null
          quantity_added?: number | null
          request_date?: string | null
          requesting_department?: string | null
          supplier?: string | null
          unit?: string | null
          updated_at?: string
          year?: string | null
        }
        Update: {
          asset_location?: string | null
          created_at?: string
          id?: string
          item_name?: string | null
          purchase_order_no?: string | null
          quantity_added?: number | null
          request_date?: string | null
          requesting_department?: string | null
          supplier?: string | null
          unit?: string | null
          updated_at?: string
          year?: string | null
        }
        Relationships: []
      }
      admin_affairs_managers: {
        Row: {
          created_at: string
          email: string
          id: string
          name: string
          permissions: string
          updated_at: string
          user_id: string | null
        }
        Insert: {
          created_at?: string
          email: string
          id?: string
          name: string
          permissions?: string
          updated_at?: string
          user_id?: string | null
        }
        Update: {
          created_at?: string
          email?: string
          id?: string
          name?: string
          permissions?: string
          updated_at?: string
          user_id?: string | null
        }
        Relationships: []
      }
      admin_affairs_purchases: {
        Row: {
          admin_request_no: string | null
          created_at: string
          executor: string | null
          id: string
          item_name: string | null
          notes: string | null
          quantity_executed: number | null
          quantity_remaining: number | null
          quantity_requested: number | null
          receipt_date: string | null
          received_by: string | null
          remarks: string | null
          request_date: string | null
          requesting_department: string | null
          system_request_no: string | null
          unit: string | null
          updated_at: string
          year: string | null
        }
        Insert: {
          admin_request_no?: string | null
          created_at?: string
          executor?: string | null
          id?: string
          item_name?: string | null
          notes?: string | null
          quantity_executed?: number | null
          quantity_remaining?: number | null
          quantity_requested?: number | null
          receipt_date?: string | null
          received_by?: string | null
          remarks?: string | null
          request_date?: string | null
          requesting_department?: string | null
          system_request_no?: string | null
          unit?: string | null
          updated_at?: string
          year?: string | null
        }
        Update: {
          admin_request_no?: string | null
          created_at?: string
          executor?: string | null
          id?: string
          item_name?: string | null
          notes?: string | null
          quantity_executed?: number | null
          quantity_remaining?: number | null
          quantity_requested?: number | null
          receipt_date?: string | null
          received_by?: string | null
          remarks?: string | null
          request_date?: string | null
          requesting_department?: string | null
          system_request_no?: string | null
          unit?: string | null
          updated_at?: string
          year?: string | null
        }
        Relationships: []
      }
      admin_affairs_summary: {
        Row: {
          created_at: string
          description: string | null
          execution_date: string | null
          id: string
          remaining_items: string | null
          remarks: string | null
          request_date: string | null
          request_number: string | null
          status: string | null
          updated_at: string
          year: string | null
        }
        Insert: {
          created_at?: string
          description?: string | null
          execution_date?: string | null
          id?: string
          remaining_items?: string | null
          remarks?: string | null
          request_date?: string | null
          request_number?: string | null
          status?: string | null
          updated_at?: string
          year?: string | null
        }
        Update: {
          created_at?: string
          description?: string | null
          execution_date?: string | null
          id?: string
          remaining_items?: string | null
          remarks?: string | null
          request_date?: string | null
          request_number?: string | null
          status?: string | null
          updated_at?: string
          year?: string | null
        }
        Relationships: []
      }
      admin_affairs_vegetables: {
        Row: {
          admin_request_no: string | null
          created_at: string
          executor: string | null
          id: string
          item_name: string | null
          notes: string | null
          quantity_executed: number | null
          quantity_remaining: number | null
          quantity_requested: number | null
          receipt_date: string | null
          received_by: string | null
          remarks: string | null
          request_date: string | null
          requesting_department: string | null
          system_request_no: string | null
          unit: string | null
          updated_at: string
          year: string | null
        }
        Insert: {
          admin_request_no?: string | null
          created_at?: string
          executor?: string | null
          id?: string
          item_name?: string | null
          notes?: string | null
          quantity_executed?: number | null
          quantity_remaining?: number | null
          quantity_requested?: number | null
          receipt_date?: string | null
          received_by?: string | null
          remarks?: string | null
          request_date?: string | null
          requesting_department?: string | null
          system_request_no?: string | null
          unit?: string | null
          updated_at?: string
          year?: string | null
        }
        Update: {
          admin_request_no?: string | null
          created_at?: string
          executor?: string | null
          id?: string
          item_name?: string | null
          notes?: string | null
          quantity_executed?: number | null
          quantity_remaining?: number | null
          quantity_requested?: number | null
          receipt_date?: string | null
          received_by?: string | null
          remarks?: string | null
          request_date?: string | null
          requesting_department?: string | null
          system_request_no?: string | null
          unit?: string | null
          updated_at?: string
          year?: string | null
        }
        Relationships: []
      }
      assets_rows: {
        Row: {
          apr: number | null
          aug: number | null
          created_at: string
          dec: number | null
          department: string | null
          description: string | null
          feb: number | null
          id: string
          item_name: string | null
          jan: number | null
          jul: number | null
          jun: number | null
          mar: number | null
          may: number | null
          nov: number | null
          oct: number | null
          price: number | null
          sep: number | null
          total_cost: number | null
          total_qty: number | null
          unit: string | null
          updated_at: string
        }
        Insert: {
          apr?: number | null
          aug?: number | null
          created_at?: string
          dec?: number | null
          department?: string | null
          description?: string | null
          feb?: number | null
          id?: string
          item_name?: string | null
          jan?: number | null
          jul?: number | null
          jun?: number | null
          mar?: number | null
          may?: number | null
          nov?: number | null
          oct?: number | null
          price?: number | null
          sep?: number | null
          total_cost?: number | null
          total_qty?: number | null
          unit?: string | null
          updated_at?: string
        }
        Update: {
          apr?: number | null
          aug?: number | null
          created_at?: string
          dec?: number | null
          department?: string | null
          description?: string | null
          feb?: number | null
          id?: string
          item_name?: string | null
          jan?: number | null
          jul?: number | null
          jun?: number | null
          mar?: number | null
          may?: number | null
          nov?: number | null
          oct?: number | null
          price?: number | null
          sep?: number | null
          total_cost?: number | null
          total_qty?: number | null
          unit?: string | null
          updated_at?: string
        }
        Relationships: []
      }
      audit_log: {
        Row: {
          action: string | null
          created_at: string
          id: string
          record_id: string | null
          table_name: string | null
          user_id: string | null
          user_name: string | null
        }
        Insert: {
          action?: string | null
          created_at?: string
          id?: string
          record_id?: string | null
          table_name?: string | null
          user_id?: string | null
          user_name?: string | null
        }
        Update: {
          action?: string | null
          created_at?: string
          id?: string
          record_id?: string | null
          table_name?: string | null
          user_id?: string | null
          user_name?: string | null
        }
        Relationships: []
      }
      budget_rows: {
        Row: {
          apr: number | null
          aug: number | null
          created_at: string
          dec: number | null
          feb: number | null
          id: string
          item: string | null
          item_code: string | null
          jan: number | null
          jul: number | null
          jun: number | null
          mar: number | null
          may: number | null
          nov: number | null
          oct: number | null
          price: number | null
          sep: number | null
          sheet: string | null
          total_cost: number | null
          total_qty: number | null
          unit: string | null
          updated_at: string
        }
        Insert: {
          apr?: number | null
          aug?: number | null
          created_at?: string
          dec?: number | null
          feb?: number | null
          id?: string
          item?: string | null
          item_code?: string | null
          jan?: number | null
          jul?: number | null
          jun?: number | null
          mar?: number | null
          may?: number | null
          nov?: number | null
          oct?: number | null
          price?: number | null
          sep?: number | null
          sheet?: string | null
          total_cost?: number | null
          total_qty?: number | null
          unit?: string | null
          updated_at?: string
        }
        Update: {
          apr?: number | null
          aug?: number | null
          created_at?: string
          dec?: number | null
          feb?: number | null
          id?: string
          item?: string | null
          item_code?: string | null
          jan?: number | null
          jul?: number | null
          jun?: number | null
          mar?: number | null
          may?: number | null
          nov?: number | null
          oct?: number | null
          price?: number | null
          sep?: number | null
          sheet?: string | null
          total_cost?: number | null
          total_qty?: number | null
          unit?: string | null
          updated_at?: string
        }
        Relationships: []
      }
      user_roles: {
        Row: {
          created_at: string
          id: string
          role: Database["public"]["Enums"]["app_role"]
          user_id: string
        }
        Insert: {
          created_at?: string
          id?: string
          role: Database["public"]["Enums"]["app_role"]
          user_id: string
        }
        Update: {
          created_at?: string
          id?: string
          role?: Database["public"]["Enums"]["app_role"]
          user_id?: string
        }
        Relationships: []
      }
    }
    Views: {
      [_ in never]: never
    }
    Functions: {
      has_role: {
        Args: {
          _role: Database["public"]["Enums"]["app_role"]
          _user_id: string
        }
        Returns: boolean
      }
      is_owner_or_manager: { Args: { _user_id: string }; Returns: boolean }
    }
    Enums: {
      app_role: "owner" | "manager" | "viewer"
    }
    CompositeTypes: {
      [_ in never]: never
    }
  }
}

type DatabaseWithoutInternals = Omit<Database, "__InternalSupabase">

type DefaultSchema = DatabaseWithoutInternals[Extract<keyof Database, "public">]

export type Tables<
  DefaultSchemaTableNameOrOptions extends
    | keyof (DefaultSchema["Tables"] & DefaultSchema["Views"])
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof (DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"] &
        DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Views"])
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? (DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"] &
      DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Views"])[TableName] extends {
      Row: infer R
    }
    ? R
    : never
  : DefaultSchemaTableNameOrOptions extends keyof (DefaultSchema["Tables"] &
        DefaultSchema["Views"])
    ? (DefaultSchema["Tables"] &
        DefaultSchema["Views"])[DefaultSchemaTableNameOrOptions] extends {
        Row: infer R
      }
      ? R
      : never
    : never

export type TablesInsert<
  DefaultSchemaTableNameOrOptions extends
    | keyof DefaultSchema["Tables"]
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"]
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"][TableName] extends {
      Insert: infer I
    }
    ? I
    : never
  : DefaultSchemaTableNameOrOptions extends keyof DefaultSchema["Tables"]
    ? DefaultSchema["Tables"][DefaultSchemaTableNameOrOptions] extends {
        Insert: infer I
      }
      ? I
      : never
    : never

export type TablesUpdate<
  DefaultSchemaTableNameOrOptions extends
    | keyof DefaultSchema["Tables"]
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"]
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"][TableName] extends {
      Update: infer U
    }
    ? U
    : never
  : DefaultSchemaTableNameOrOptions extends keyof DefaultSchema["Tables"]
    ? DefaultSchema["Tables"][DefaultSchemaTableNameOrOptions] extends {
        Update: infer U
      }
      ? U
      : never
    : never

export type Enums<
  DefaultSchemaEnumNameOrOptions extends
    | keyof DefaultSchema["Enums"]
    | { schema: keyof DatabaseWithoutInternals },
  EnumName extends DefaultSchemaEnumNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaEnumNameOrOptions["schema"]]["Enums"]
    : never = never,
> = DefaultSchemaEnumNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[DefaultSchemaEnumNameOrOptions["schema"]]["Enums"][EnumName]
  : DefaultSchemaEnumNameOrOptions extends keyof DefaultSchema["Enums"]
    ? DefaultSchema["Enums"][DefaultSchemaEnumNameOrOptions]
    : never

export type CompositeTypes<
  PublicCompositeTypeNameOrOptions extends
    | keyof DefaultSchema["CompositeTypes"]
    | { schema: keyof DatabaseWithoutInternals },
  CompositeTypeName extends PublicCompositeTypeNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[PublicCompositeTypeNameOrOptions["schema"]]["CompositeTypes"]
    : never = never,
> = PublicCompositeTypeNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[PublicCompositeTypeNameOrOptions["schema"]]["CompositeTypes"][CompositeTypeName]
  : PublicCompositeTypeNameOrOptions extends keyof DefaultSchema["CompositeTypes"]
    ? DefaultSchema["CompositeTypes"][PublicCompositeTypeNameOrOptions]
    : never

export const Constants = {
  public: {
    Enums: {
      app_role: ["owner", "manager", "viewer"],
    },
  },
} as const
