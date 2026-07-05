import { createFileRoute } from "@tanstack/react-router";
import AdminAffairsSystem from "@/App";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Linah Farms - نظام إدارة الشؤون" },
      { name: "description", content: "نظام إدارة الشؤون الإدارية لمزارع لينة" },
    ],
  }),
  component: AdminAffairsSystem,
});
