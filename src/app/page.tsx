import { ExtractorApp } from "@/components/ExtractorApp";
import { InlineJeProbe } from "@/components/InlineJeProbe";

export const dynamic = "force-dynamic";
export const revalidate = 0;

export default function Home() {
  return (
    <>
      <InlineJeProbe pageName="信息提取" />
      <ExtractorApp />
    </>
  );
}
